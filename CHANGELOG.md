# Changelog

## [0.2.0] - 2026-09-04

### Added

- Bounded `LayoutOptimizer` for automatic DOCX density search.
- Per-iteration parameter history in `optimization-history.json`.
- Visual diagnostics with page overlays and diff images.
- Debug PDF for quick visual inspection.
- Report asset endpoint for diagnostic artifacts.
- Explicit application version reporting (`0.2.0`).

### Changed

- Conversion pipeline now selects the best candidate before producing the final DOCX.
- Visual comparison is now also used as an optimization feedback signal.
- README updated to document the v0.2 architecture and diagnostics.

### Compatibility

- Existing `/api/v1/convert`, file download, and JSON report endpoints remain available.
