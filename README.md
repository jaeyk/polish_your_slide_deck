# Polish Your Slide Deck

An MCP server that reads PowerPoint decks (preferred) or PDFs from `data/`, exports their text as editable Markdown, and writes polished, editable PPTX copies to `output/`. It offers two audience profiles: `academic` (default) and `professional`.

## Setup

Use Python 3.10 or newer:

```sh
python -m pip install -r requirements.txt
```

Put `.pptx` files in `data/`. The project MCP configuration is in `.vscode/mcp.json`; start or restart the servers from VS Code's MCP controls. Figma MCP is optional and will prompt you to authenticate with Figma.

## Workflow

1. `list_decks` finds `.pptx` and `.pdf` files in `data/`.
2. `export_deck_to_markdown(filename)` writes `output/<deck-name>.md`. Edit text under the slide and shape markers; leave the marker comments unchanged.
3. `apply_markdown_edits(filename)` writes `output/<deck-name>-edited.pptx`.
4. `polish_deck(filename, audience="academic", ...)` applies Markdown edits and audience styling to `output/<deck-name>-polished.pptx`. Without a Markdown file, it uses an existing `-edited.pptx` copy if available.

Choose `audience="academic"` for a restrained serif treatment or `audience="professional"` for a sans-serif treatment. Optional `primary_color`, `accent_color` (hex colors), and `font_family` parameters let an agent pass selected values from a Figma design system into the formatter.

For PPTX input, the round trip preserves existing images, charts, and geometry while updating text. For PDF input, selectable text is rebuilt as an editable slide per PDF page; original layout and graphics are not preserved, and scanned PDFs need OCR first.

Figma MCP is a separate, optional MCP connection configured in `.vscode/mcp.json`. It can provide design context to the agent, but this server does not access Figma files or credentials directly. The agent can pass Figma-derived font and color tokens to `polish_deck`; it does not automatically convert Figma components into PowerPoint elements.

## Existing Project Considered

[office-powerpoint-mcp-server](https://github.com/skittermonster/office-powerpoint-mcp-server) has useful theme-application and inspection capabilities. This project uses the same established `python-pptx` and MCP SDK foundations, while keeping the two-profile, folder-based workflow small and explicit instead of taking on that server's broader Office automation surface.
