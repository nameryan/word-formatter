# Repository Guidelines

## Project Structure & Module Organization

The app is a browser-based `.docx` formatter. The production entry point is `index.html`, which includes the runtime UI and the inlined formatter logic used under `file://`. Shared styling lives in `css/style.css`. The `js/` directory contains the modular reference source (`app.js`, `doc-patcher.js`, `style-patcher.js`, `xml-parser.js`, `docx-reader.js`, `docx-writer.js`) and should stay behaviorally aligned with the inlined script. `example/` is reserved for sample assets. The `leak-check/` folder is a separate Python service and should be treated as an isolated subproject.

## Build, Test, and Development Commands

- `open index.html` — launch the formatter directly in a browser on macOS.
- `python3 -m http.server` — optional local server for testing module-based loading instead of `file://`.
- `node --check js/app.js` — syntax-check a JS module; repeat for touched files.
- `python3 - <<'PY' ...` to extract the inline script from `index.html`, then `node --check /tmp/index-inline.js` — useful when validating the embedded runtime script after edits.

There is no build pipeline or package manager in the main app; keep the workflow lightweight.

## Coding Style & Naming Conventions

Use 2-space indentation in HTML, CSS, and inline JavaScript. Prefer ASCII in code unless the file already contains Chinese UI text or OOXML-specific values. Keep helper names descriptive (`patchDocument`, `getParaStyleId`, `applyRunFormatting`). Preserve the repository pattern: XML helpers are small and namespace-safe, document mutations are explicit, and UI state lives near DOM wiring. When changing formatter behavior, update both `index.html` and the matching file in `js/`.

## Testing Guidelines

There is no automated test suite yet, so rely on focused regression checks. Validate syntax with `node --check` for edited JS files and re-test with representative `.docx` files: cover page, TOC, heading numbering, tables, and page-margin cases. Confirm that cover pages remain untouched when enabled, headings do not double-number, and table content keeps structure.

## Commit & Pull Request Guidelines

Follow the existing history style: concise, imperative commit subjects such as `Fix indent...`, `Add heading auto-numbering...`, or `Initial release...`. Keep each commit scoped to one behavior change. PRs should include: a short problem statement, the affected document patterns, manual test notes, and before/after screenshots when UI behavior changes.

## DOCX Regression Checklist

Before merging formatter changes, run at least one manual pass covering these cases:

- Cover page enabled: page 1 typography stays unchanged while body pages are reformatted.
- TOC or preface pages: no stray numbering like `0.0.1`, and TOC entries are not re-indented incorrectly.
- Heading numbering: existing numbered headings do not become double-numbered.
- Tables and captions: table structure stays intact and special styles are not given body indents.
- Page layout: margin changes apply as expected; if editing section logic, test a multi-section `.docx`.
