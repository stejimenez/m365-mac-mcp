"""
PowerPoint Controller for Microsoft 365 for Mac MCP (m365-mac-mcp).
Handles all PowerPoint automation operations on macOS.
"""

import asyncio
import uuid
import tempfile
from pathlib import Path
from typing import (
    Any,
    Dict,
    List,
    Optional,
)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

import re
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from integrations.applescript_bridge import AppleScriptBridge
from utils.config import open_office_template, resolve_template_path
from utils.logger import setup_logger
from utils.validators import atomic_save_via_temp, resolve_file_path

logger = setup_logger(__name__)

class PowerPointController:
    """Controller for PowerPoint operations using both AppleScript and python-pptx."""

    _LAYOUT_MAP = {
        "title slide": 0,
        "title_slide": 0,
        "title": 0,
        "title and content": 1,
        "title_and_content": 1,
        "content": 1,
        "bullets": 1,
        "section header": 2,
        "section_header": 2,
        "section": 2,
        "two content": 3,
        "two_content": 3,
        "comparison": 4,
        "title only": 5,
        "title_only": 5,
        "blank": 6,
        "content with caption": 7,
        "content_with_caption": 7,
        "picture with caption": 8,
        "picture_with_caption": 8,
    }

    def __init__(self, max_presentations: int = 20):
        self.applescript = AppleScriptBridge()
        self.active_presentations: Dict[str, Dict[str, Any]] = {}
        self.max_presentations = max_presentations
        # Use system temp directory with a subdirectory
        self.temp_dir = Path(tempfile.gettempdir()) / "m365-mac-mcp"
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    def _evict_if_needed(self) -> None:
        """Keep active presentation count bounded, preferring clean presentations first."""
        while len(self.active_presentations) >= self.max_presentations:
            evict_id = next(
                (pid for pid, data in self.active_presentations.items() if not data.get("dirty")),
                None,
            )
            if evict_id is None:
                evict_id = next(iter(self.active_presentations))
            self.active_presentations.pop(evict_id, None)
            logger.info(f"Evicted in-memory PowerPoint presentation {evict_id} to stay within limit ({self.max_presentations})")

    def _resolve_presentation_id(self, presentation_id: Optional[str]) -> str:
        """Resolve an optional presentation_id (auto-selects when exactly 1 presentation is active)."""
        if presentation_id and str(presentation_id).strip():
            pid = str(presentation_id).strip()
            if pid not in self.active_presentations:
                raise ValueError(f"Presentation {pid} not found")
            return pid
        if len(self.active_presentations) == 1:
            return next(iter(self.active_presentations))
        if len(self.active_presentations) == 0:
            raise ValueError("No active presentation found. Call create_presentation first.")
        raise ValueError(
            f"Multiple presentations are active ({len(self.active_presentations)}). Pass presentation_id explicitly."
        )

    def _resolve_layout_index(self, prs: Presentation, layout: str) -> int:
        key = str(layout or "Title and Content").strip().lower()
        idx = self._LAYOUT_MAP.get(key, 1)
        if idx >= len(prs.slide_layouts):
            return min(1, len(prs.slide_layouts) - 1)
        return idx

    @staticmethod
    def _clean_bullet_lines(text: str) -> str:
        """Strip redundant leading bullet glyphs so PowerPoint placeholders never render double bullets."""
        return "\n".join(
            re.sub(r"^\s*(?:[•●◦▪▫–—\-*]|\d+\.)\s+", "", line)
            if re.match(r"^\s*[•●◦▪▫\-*]\s+", line)
            else line
            for line in text.splitlines()
        )

    def _find_slide(self, slide_id: Optional[str]):
        """Find (presentation_id, slide_id, slide_obj) by slide_id, or latest slide if omitted with 1 active deck."""
        if slide_id and str(slide_id).strip():
            sid = str(slide_id).strip()
            for pres_id, pres_data in self.active_presentations.items():
                if sid in pres_data["slides"]:
                    return pres_id, sid, pres_data["slides"][sid]["slide_object"]
            raise ValueError(f"Slide {sid} not found")
        if len(self.active_presentations) == 1:
            pres_id = next(iter(self.active_presentations))
            slides_dict = self.active_presentations[pres_id]["slides"]
            if slides_dict:
                sid = list(slides_dict.keys())[-1]
                return pres_id, sid, slides_dict[sid]["slide_object"]
        raise ValueError("Slide not found. Provide slide_id explicitly.")

    async def create_presentation(
        self,
        title: str,
        theme: str = "default",
        template_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a new PowerPoint presentation.
        
        Args:
            title: Presentation title
            theme: Theme name or path
            template_path: Optional template file path
            
        Returns:
            Dict with presentation metadata
        """
        try:
            self._evict_if_needed()
            presentation_id = str(uuid.uuid4())
            
            # Create presentation using python-pptx (explicit template -> M365_MCP_PPT_TEMPLATE -> 16:9 blank)
            resolved_template = resolve_template_path(
                template_path,
                default_env_var="M365_MCP_PPT_TEMPLATE",
                allowed_extensions=(".pptx", ".potx"),
            )
            if resolved_template is not None:
                prs = Presentation(open_office_template(resolved_template))
            else:
                prs = Presentation()
                # Modern 16:9 widescreen default instead of python-pptx legacy 4:3
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
            
            # Set title if first slide exists, otherwise add Title Slide
            if prs.slides:
                title_slide = prs.slides[0]
                if title_slide.shapes.title:
                    title_slide.shapes.title.text = title
            else:
                title_slide_layout = prs.slide_layouts[0]  # Title slide layout
                title_slide = prs.slides.add_slide(title_slide_layout)
                if title_slide.shapes.title:
                    title_slide.shapes.title.text = title
            
            # Track temp path (deferred disk write until save_presentation)
            temp_file = self.temp_dir / f"{presentation_id}.pptx"
            
            # Register Slide 1 so callers can immediately add subtitle/text/notes to it
            now_ts = asyncio.get_event_loop().time()
            first_slide_id = str(uuid.uuid4())
            first_slide_meta = {
                "slide_id": first_slide_id,
                "presentation_id": presentation_id,
                "layout": "Title Slide",
                "index": 0,
                "created_at": now_ts,
            }
            
            # Store presentation metadata
            presentation_data = {
                "presentation_id": presentation_id,
                "title": title,
                "theme": theme,
                "file_path": str(temp_file),
                "slide_count": len(prs.slides),
                "first_slide_id": first_slide_id,
                "slides": [{"slide_id": first_slide_id, "index": 0, "layout": "Title Slide"}],
                "created_at": now_ts,
            }
            if resolved_template is not None:
                presentation_data["template_used"] = str(resolved_template)
            
            self.active_presentations[presentation_id] = {
                "metadata": presentation_data,
                "pptx_object": prs,
                "dirty": True,
                "slides": {
                    first_slide_id: {
                        "metadata": first_slide_meta,
                        "slide_object": title_slide,
                    }
                },
            }
            
            logger.info(f"Created presentation: {title} ({presentation_id})")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Failed to create presentation: {e}")
            raise
    
    async def add_slide(
        self,
        presentation_id: Optional[str] = None,
        layout: str = "Title and Content",
        position: Optional[int] = None
    ) -> Dict[str, Any]:
        """Add a new slide to a presentation."""
        try:
            presentation_id = self._resolve_presentation_id(presentation_id)
            prs = self.active_presentations[presentation_id]["pptx_object"]
            slide_id = str(uuid.uuid4())
            
            layout_index = self._resolve_layout_index(prs, layout)
            slide_layout = prs.slide_layouts[layout_index]
            slide = prs.slides.add_slide(slide_layout)
            slide_index = len(prs.slides) - 1
            
            # Store slide metadata
            slide_data = {
                "slide_id": slide_id,
                "presentation_id": presentation_id,
                "layout": layout,
                "index": slide_index,
                "created_at": asyncio.get_event_loop().time()
            }
            
            self.active_presentations[presentation_id]["slides"][slide_id] = {
                "metadata": slide_data,
                "slide_object": slide
            }
            self.active_presentations[presentation_id]["dirty"] = True
            
            # Update presentation metadata
            self.active_presentations[presentation_id]["metadata"]["slide_count"] = len(prs.slides)
            self.active_presentations[presentation_id]["metadata"].setdefault("slides", []).append({
                "slide_id": slide_id,
                "index": slide_index,
                "layout": layout,
            })
            
            logger.info(f"Added slide to presentation {presentation_id}")
            return slide_data
            
        except Exception as e:
            logger.error(f"Failed to add slide: {e}")
            raise

    async def batch_add_slides(
        self,
        presentation_id: Optional[str] = None,
        slides: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Any]:
        """Add and populate multiple slides (title, subtitle, bullets/content, speaker notes) in ONE atomic call."""
        try:
            presentation_id = self._resolve_presentation_id(presentation_id)
            if not slides or not isinstance(slides, list):
                raise ValueError("slides must be a non-empty list of slide specification dicts.")

            entry = self.active_presentations[presentation_id]
            prs = entry["pptx_object"]
            now_ts = asyncio.get_event_loop().time()
            added_summaries: List[Dict[str, Any]] = []

            for spec in slides:
                if not isinstance(spec, dict):
                    raise ValueError("Each slide entry in slides must be a dict.")
                layout_name = str(spec.get("layout") or "Title and Content")
                layout_index = self._resolve_layout_index(prs, layout_name)
                slide_obj = prs.slides.add_slide(prs.slide_layouts[layout_index])
                slide_id = str(uuid.uuid4())
                slide_index = len(prs.slides) - 1

                title_text = spec.get("title")
                title_shape = slide_obj.shapes.title
                if title_text is not None and title_shape is not None:
                    title_shape.text_frame.text = str(title_text)
                    if spec.get("title_formatting"):
                        await self._apply_text_formatting(title_shape.text_frame, spec["title_formatting"])

                body_text: Optional[str] = None
                if spec.get("bullets") is not None:
                    raw_bullets = spec["bullets"]
                    if isinstance(raw_bullets, list):
                        body_text = "\n".join(str(b) for b in raw_bullets)
                    else:
                        body_text = str(raw_bullets)
                    body_text = self._clean_bullet_lines(body_text)
                elif spec.get("content") is not None:
                    body_text = self._clean_bullet_lines(str(spec["content"]))
                elif spec.get("subtitle") is not None:
                    body_text = str(spec["subtitle"])

                if body_text is not None:
                    target_placeholder = None
                    for shape in slide_obj.placeholders:
                        if title_shape is not None and shape == title_shape:
                            continue
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            target_placeholder = shape
                            break
                    if target_placeholder is not None:
                        tf = target_placeholder.text_frame
                        tf.text = body_text
                    else:
                        tb = slide_obj.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11.3), Inches(4.8))
                        tf = tb.text_frame
                        tf.text = body_text
                    if spec.get("content_formatting"):
                        await self._apply_text_formatting(tf, spec["content_formatting"])

                if spec.get("notes"):
                    slide_obj.notes_slide.notes_text_frame.text = str(spec["notes"])

                slide_meta = {
                    "slide_id": slide_id,
                    "presentation_id": presentation_id,
                    "layout": layout_name,
                    "index": slide_index,
                    "created_at": now_ts,
                }
                entry["slides"][slide_id] = {"metadata": slide_meta, "slide_object": slide_obj}
                summary = {"slide_id": slide_id, "index": slide_index, "layout": layout_name}
                if title_text is not None:
                    summary["title"] = str(title_text)
                entry["metadata"].setdefault("slides", []).append(summary)
                added_summaries.append(summary)

            entry["dirty"] = True
            entry["metadata"]["slide_count"] = len(prs.slides)
            logger.info(f"Batch-added {len(added_summaries)} slides to presentation {presentation_id}")
            return {
                "status": "success",
                "presentation_id": presentation_id,
                "slides_added": len(added_summaries),
                "total_slides": len(prs.slides),
                "slides": added_summaries,
            }
        except Exception as e:
            logger.error(f"Failed to batch-add slides: {e}")
            raise
    
    async def add_text(
        self,
        slide_id: Optional[str] = None,
        text: str = "",
        placeholder: str = "content",
        formatting: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Add text to a slide."""
        try:
            presentation_id, slide_id, slide_obj = self._find_slide(slide_id)
            
            # Add text based on placeholder type (never let 'content'/'subtitle' overwrite title!)
            ph = (placeholder or "content").strip().lower()
            title_shape = slide_obj.shapes.title
            if ph == "title" and title_shape:
                text_frame = title_shape.text_frame
                text_frame.text = text
            elif ph in {"content", "body", "subtitle"}:
                target_placeholder = None
                for shape in slide_obj.placeholders:
                    if title_shape is not None and shape == title_shape:
                        continue
                    if hasattr(shape, "text_frame") and shape.text_frame:
                        target_placeholder = shape
                        break
                if target_placeholder:
                    cleaned_text = self._clean_bullet_lines(text) if ph in {"content", "body"} else text
                    text_frame = target_placeholder.text_frame
                    text_frame.text = cleaned_text
                else:
                    left = Inches(1)
                    top = Inches(1.8 if ph != "subtitle" else 3.8)
                    width = Inches(8)
                    height = Inches(4.5 if ph != "subtitle" else 1.2)
                    textbox = slide_obj.shapes.add_textbox(left, top, width, height)
                    text_frame = textbox.text_frame
                    text_frame.text = text
            else:
                left = Inches(1)
                top = Inches(1.8)
                width = Inches(8)
                height = Inches(4.5)
                textbox = slide_obj.shapes.add_textbox(left, top, width, height)
                text_frame = textbox.text_frame
                text_frame.text = text
            
            # Apply formatting if provided
            if formatting and text_frame:
                await self._apply_text_formatting(text_frame, formatting)
            
            self.active_presentations[presentation_id]["dirty"] = True
            logger.info(f"Added text to slide {slide_id}")
            return {
                "status": "success",
                "slide_id": slide_id,
                "text_length": len(text),
                "placeholder": placeholder
            }
            
        except Exception as e:
            logger.error(f"Failed to add text: {e}")
            raise
    
    async def add_image(
        self,
        slide_id: Optional[str] = None,
        image_source: str = "",
        position: Optional[Dict[str, float]] = None,
        size: Optional[Dict[str, float]] = None
    ) -> Dict[str, Any]:
        """Add an image to a slide."""
        try:
            presentation_id, slide_id, slide_obj = self._find_slide(slide_id)
            position = position or {"x": 1.0, "y": 1.0}
            size = size or {"width": 4.0, "height": 3.0}
            
            # Handle image source
            if image_source.startswith(("http://", "https://")):
                raise NotImplementedError("URL image download not yet implemented")
            else:
                image_path = Path(image_source).expanduser().resolve()
                if not image_path.exists():
                    raise FileNotFoundError(f"Image file not found: {image_source}")
            
            left = Inches(position.get("x", 1))
            top = Inches(position.get("y", 1))
            width = Inches(size.get("width", 4))
            height = Inches(size.get("height", 3))
            
            slide_obj.shapes.add_picture(str(image_path), left, top, width, height)
            self.active_presentations[presentation_id]["dirty"] = True
            
            logger.info(f"Added image to slide {slide_id}")
            return {
                "status": "success",
                "slide_id": slide_id,
                "image_source": image_source,
                "position": position,
                "size": size
            }
            
        except Exception as e:
            logger.error(f"Failed to add image: {e}")
            raise
    
    async def add_speaker_notes(
        self,
        slide_id: Optional[str] = None,
        notes: str = ""
    ) -> Dict[str, Any]:
        """Add speaker notes to a slide."""
        try:
            presentation_id, slide_id, slide_obj = self._find_slide(slide_id)
            notes_slide = slide_obj.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = notes
            
            self.active_presentations[presentation_id]["dirty"] = True
            logger.info(f"Added speaker notes to slide {slide_id}")
            return {
                "status": "success",
                "slide_id": slide_id,
                "notes_length": len(notes)
            }
            
        except Exception as e:
            logger.error(f"Failed to add speaker notes: {e}")
            raise
    
    async def save_presentation(
        self,
        presentation_id: Optional[str] = None,
        file_path: str = "",
        format: str = "pptx"
    ) -> Dict[str, Any]:
        """Save a presentation to file."""
        try:
            presentation_id = self._resolve_presentation_id(presentation_id)
            prs = self.active_presentations[presentation_id]["pptx_object"]
            save_path = resolve_file_path(file_path, must_exist=False, for_write=True)
            
            save_path.parent.mkdir(parents=True, exist_ok=True)
            
            if format.lower() == "pptx":
                await asyncio.to_thread(atomic_save_via_temp, prs.save, save_path)
            elif format.lower() == "pdf":
                pptx_path = save_path.with_suffix(".pptx")
                await asyncio.to_thread(atomic_save_via_temp, prs.save, pptx_path)
                logger.warning(f"PDF export not implemented, saved as PPTX: {pptx_path}")
                save_path = pptx_path
            else:
                raise ValueError(f"Unsupported format: {format}")
            
            self.active_presentations[presentation_id]["metadata"]["file_path"] = str(save_path)
            self.active_presentations[presentation_id]["dirty"] = False
            
            logger.info(f"Saved presentation to {save_path}")
            return {
                "status": "success",
                "presentation_id": presentation_id,
                "file_path": str(save_path),
                "format": format
            }
            
        except Exception as e:
            logger.error(f"Failed to save presentation: {e}")
            raise
    
    async def _apply_text_formatting(
        self,
        text_frame,
        formatting: Dict[str, Any]
    ) -> None:
        """Apply text formatting to a text frame.
        
        Args:
            text_frame: PowerPoint text frame object
            formatting: Formatting options
        """
        try:
            align_map = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY,
            }
            color_val = formatting.get("color") or formatting.get("font_color")
            rgb_color = None
            if isinstance(color_val, str) and color_val.lstrip("#") and len(color_val.lstrip("#")) == 6:
                hex_color = color_val.lstrip("#")
                rgb_color = RGBColor(*(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

            for paragraph in text_frame.paragraphs:
                if "font_size" in formatting:
                    paragraph.font.size = Pt(formatting["font_size"])
                if "font_name" in formatting:
                    paragraph.font.name = formatting["font_name"]
                if "bold" in formatting:
                    paragraph.font.bold = formatting["bold"]
                if "italic" in formatting:
                    paragraph.font.italic = formatting["italic"]
                if rgb_color is not None:
                    paragraph.font.color.rgb = rgb_color
                if "alignment" in formatting and formatting["alignment"] in align_map:
                    paragraph.alignment = align_map[formatting["alignment"]]
            
        except Exception as e:
            logger.warning(f"Failed to apply some text formatting: {e}")
    
    async def get_presentation_info(self, presentation_id: Optional[str] = None) -> Dict[str, Any]:
        """Get information about a presentation."""
        presentation_id = self._resolve_presentation_id(presentation_id)
        return self.active_presentations[presentation_id]["metadata"]
    
    async def list_presentations(self) -> List[Dict[str, Any]]:
        """List all active presentations.
        
        Returns:
            List of presentation metadata
        """
        return [data["metadata"] for data in self.active_presentations.values()]