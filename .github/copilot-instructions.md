## Project Context
- This is a **pure Dart package** (`excel_plus`) for reading/writing `.xlsx` files.
- NOT a Flutter app. No UI, no state management, no networking.
- Must work on all platforms: VM, Web, mobile (via Flutter dependencies).
- Uses `part of excel_plus;` library system across all source files.

---

## 1. Architecture (Layered Package Structure)
- Follow layered architecture within `lib/src/`:

```
lib/src/
  core/       → Main Excel class, config/constants
  models/     → Data classes, enums, value types (CellValue, CellStyle, etc.)
  sheet/      → Sheet operations and manipulation
  reader/     → .xlsx parsing logic
  writer/     → .xlsx save/export logic
  utils/      → Helpers (archive, cell utilities, collections)
  platform/   → Conditional imports (web vs stub)
```

- All files use `part of excel_plus;` — the library entry point is `lib/excel_plus.dart`.
- Keep layers separate: reader should not depend on writer logic and vice versa.

---

## 2. Performance & Memory
- Prioritize memory efficiency — package must handle large files (100k+ rows) without crashing.
- Prefer streaming/SAX parsing (`parseEvents()`) over full DOM parsing where possible.
- Avoid storing duplicate data (e.g., same string in multiple structures).
- Use lazy loading for sheet data when feasible.
- Profile before optimizing — focus on measurable bottlenecks.

---

## 3. Naming Conventions
- Variables: camelCase
- Files: snake_case.dart
- Classes: PascalCase
- Constants: UPPER_SNAKE_CASE
- Private members: prefix with `_`

---

## 4. Code Style
- Avoid unnecessary comments.
- Only add comments for complex logic, important notes, or public API docs.
- Keep code clean and readable.
- Avoid over-engineering.
- Avoid excessive use of `dynamic`, but allowed when necessary.
- Public API classes and methods should have dartdoc comments (`///`).

---

## 5. Dependencies
- Keep dependencies minimal — this is a library, not an app.
- Current deps: `archive`, `xml`, `collection`, `equatable`, `web`.
- Add packages using terminal: `dart pub add <package_name>`
- Do NOT manually edit pubspec.yaml to add dependencies.

---

## 6. Error Handling
- Use structured failure/success patterns.
- Throw typed exceptions for parse errors and invalid data.
- Never silently swallow errors — at minimum log them.
- Validate input at public API boundaries only.

---

## 7. Modularity & File Size
- Keep code modular.
- Avoid large files.
- Maximum file length: 400–500 lines.
- Split large classes across multiple files if needed (using `part of`).

---

## 8. Testing
- All public API changes must have corresponding tests.
- Test with real `.xlsx` files in `test/test_resources/`.
- Cover both read and write round-trips.
- Run `dart test` before committing.
- Run `dart analyze` after any code change and fix all issues before proceeding.
- Commit after every meaningful change with a clear, descriptive message.

---

## 9. Platform Compatibility
- Package must compile on all Dart platforms (VM, Web, AOT).
- Use conditional imports for platform-specific code (`platform/` folder).
- Never use `dart:io` directly in library code — isolate behind conditional imports.

---

## 10. General Rules
- Do not mix architecture styles.
- Avoid hardcoded values — use constants from `core/config.dart`.
- Remove unused code.
- Keep code production-ready and scalable.
- Maintain backward compatibility for public API changes.