# Changelog

## 1.0.2 - Footer alignment

- Center-aligned the footer safety note and copyright link across the full application width.
- No extraction logic, pagination logic, conversion behavior, ZIP structure, or bilingual functionality was changed.

## 1.0.1 - Build fix

- Added explicit `System.IO` imports to all source files that use `File`, `Path`, `Directory`, `FileStream`, `MemoryStream`, `StreamWriter`, `FileMode`, `FileAccess`, `FileShare`, or `FileOptions`.
- Fixed the nullable-file warning in `CanExtractSelectedFile()` by snapshotting the selected path into a local non-null variable before validation.
- No extraction behavior, UI layout, bilingual toggle, image ordering, pagination logic, or ZIP structure was intentionally changed.

## 1.0.0

- Initial release.
