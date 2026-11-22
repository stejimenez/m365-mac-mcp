# Repository Agent Rules (`m365-mac-mcp`)

## GitHub Release & Semantic Versioning Policy

Whenever updating GitHub for this repository, **always** cut a versioned release using industry-standard **Semantic Versioning (`SemVer 2.0.0` — `vMAJOR.MINOR.PATCH`)**:

- **`vX.Y.Z` — PATCH (`Z+1`)**: Bug fixes, minor hardening, test additions, or documentation updates.
- **`vX.Y.0` — MINOR (`Y+1`, reset `Z=0`)**: Backward-compatible new features, MCP tools, profiles, or skills.
- **`vX.0.0` — MAJOR (`X+1`, reset `Y=0, Z=0`)**: Breaking API, tool-schema, environment-variable, or architecture changes.

### Mandatory Checklist on Every GitHub Push

1. **Determine Version Bump**: Compare changes since the latest Git tag (`git describe --tags --abbrev=0`) and select `PATCH`, `MINOR`, or `MAJOR`.
2. **Sync All Version References**:
   - `src/__init__.py` → `__version__`
   - `src/m365_mac_mcp_server.py` → `server_version`
   - `README.md` → version badge
   - `CHANGELOG.md` → dated `## [X.Y.Z] - YYYY-MM-DD` release entry
3. **Verify Before Push**: Run `.venv/bin/pytest tests/test_basic.py tests/test_word_excel_rw.py -v` and confirm all tests pass.
4. **Commit, Tag & Push**:
   - Commit the release changes cleanly.
   - Create an annotated tag: `git tag -a vX.Y.Z -m "Release vX.Y.Z: <summary>"`
   - Push commit and tag together: `git push origin main --follow-tags`
5. **Publish GitHub Release**:
   - Create the matching GitHub Release automatically via `/opt/homebrew/bin/gh release create vX.Y.Z --title "vX.Y.Z — <Short Title>" --notes "<Release Notes>"`.

> **Note**: Do not commit, tag, push, or publish a release until the user explicitly requests pushing/updating GitHub.
