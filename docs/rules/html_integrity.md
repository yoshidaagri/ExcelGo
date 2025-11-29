# HTML Integrity Check

When editing HTML files, especially `index.html`, you MUST verify the file structure after any modification.

## Rules

1.  **No Duplication**: Ensure that the `<html>`, `<head>`, and `<body>` tags are not duplicated. The file must have exactly one of each.
2.  **Complete Structure**: Verify that all sections (header, main, scripts) are present and correctly nested.
3.  **Post-Edit Verification**: After using `replace_file_content` or `multi_replace_file_content` on an HTML file, ALWAYS read the file back (using `view_file`) or use `grep_search` to check for multiple `<!DOCTYPE html>` or `<html>` tags.
4.  **Atomic Writes**: If a file becomes corrupted or duplicated, use `write_to_file` to overwrite it completely with the correct content, rather than trying to patch it with further replacements.

## Verification Command

Run this command to check for duplication:
```bash
grep -c "<html" path/to/index.html
```
If the count is greater than 1, the file is corrupted.
