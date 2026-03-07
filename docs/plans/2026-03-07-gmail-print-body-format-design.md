# Gmail Print Body Format Design

**Problem**

The current PDF flow renders the raw `text/html` email body inside a custom wrapper. That keeps more of the sender's original HTML, but it does not resemble Gmail's print PDF and can inherit aggressive Outlook/Word layout rules that push the real content to later pages.

**Goal**

Add a selectable body output mode so the user can choose between:

- `PDF estilo Gmail`
- `HTML original`
- `Ambos`

The Gmail PDF mode should produce a stable, Gmail-like print layout. The original HTML mode should preserve the source body for audit/debugging.

**Constraints**

- Keep the existing IMAP download flow.
- Preserve inline `cid:` images in both output modes.
- Do not fail the whole email download when Gmail-style PDF rendering hits malformed HTML.
- Keep Playwright in the isolated worker process.

**Proposed Architecture**

1. Introduce a rendering helper module dedicated to email body preparation.
2. Split body handling into two artifacts:
   - Original HTML artifact: mostly untouched source HTML, wrapped only when needed to open as a standalone file.
   - Gmail-print artifact: sanitized body fragment mounted into a controlled Gmail-style print template.
3. Keep attachment extraction and manifest logic unchanged.
4. Only start the Playwright worker when the selected output mode needs a PDF.

**Rendering Strategy**

- Extract the useful body fragment from the email HTML.
- Remove document-level shells and unsafe print/layout rules (`html`, `head`, `body`, scripts, Outlook/Word styles, large margins, absolute positioning, page-break noise).
- Preserve safe content structure (tables, images, links, text formatting, blockquotes).
- Apply a fixed Gmail-style print shell for title, account, sender, recipients, date, and body.

**UI Changes**

- Keep `Descargar cuerpo` as the main toggle.
- Add a new `Formato del cuerpo` selector with the three approved options.
- Persist the selection in `settings.json`.
- Disable the selector when body download is off.

**Fallback Rules**

- `HTML original`: always save `Mensaje_Original.html`.
- `PDF estilo Gmail`: save `Mensaje_Gmail.pdf`.
- `Ambos`: save both files.
- If Gmail PDF generation fails, save `Mensaje_Gmail.html` as a readable fallback so the body is not lost.

**Testing Strategy**

- Add unit tests for HTML sanitization and template generation.
- Verify that inline images survive sanitization.
- Verify that document wrappers and unsafe styles are removed from Gmail PDF mode.
- Run a real Playwright PDF generation check after implementation.
