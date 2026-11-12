# Polish Your Slide Deck

An MCP server that reads PowerPoint decks (preferred) or PDFs from `data/`, exports their text as editable Markdown, and writes polished, editable PPTX copies to `output/`. It offers four style profiles: `newsroom_light` (default), `civic_signal`, `academic`, and `professional`.

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
4. `polish_deck(filename, audience="newsroom_light", ...)` applies Markdown edits and the style profile to `output/<deck-name>-polished.pptx`. Without a Markdown file, it uses an existing `-edited.pptx` copy if available.
5. `get_design_preferences()` returns [design-preferences.md](design-preferences.md), the style rules an agent should read before proposing or applying a style.

### Style profiles

- **`newsroom_light` (default):** cool white background (#F5F7FA), near-black Overpass text, and Carolina Navy (#13294B) emphasis. Visible text highlights become navy blocks with light text, and shape outlines turn navy. White screenshots and charts blend into the background.
- **`civic_signal`:** the dark alternative, with a deep navy background (#0B1A2E), white Public Sans text, and signal-yellow (#FFD34E) emphasis. Visible text highlights become yellow blocks with navy text, white highlights (invisible on the original light slides) are removed, and shape outlines turn yellow. Any text color below 7:1 contrast is replaced. On this dark background, pictures that are more than 10% transparent, such as charts exported without a background, get a white plate behind them so their dark lines stay visible.
- **`academic`:** a restrained Georgia serif treatment on the deck's own background.
- **`professional`:** an Aptos sans-serif treatment on the deck's own background.

Every profile raises text to the large-type floors (titles 40 pt, body 24 pt, sources, notes, and slide numbers 18 pt). Title placeholders of 12 or more words are treated as captions and get the body floor. Text is never shrunk. Text-only slides are vertically centered, sitting slightly above the middle; slides with pictures or drawn shapes keep their layout. When enlarged text may run off the slide or into a picture, `polish_deck` lists those slides so you can shorten the text through the Markdown round trip.

`newsroom_light` uses the free Overpass font (`brew install --cask font-overpass`). `civic_signal` uses Public Sans (`brew install --cask font-public-sans`). Afterwards, remove the variable `PublicSans[wght].ttf` and `PublicSans-Italic[wght].ttf` files from `~/Library/Fonts`. Their default instance is Thin, which apps may use in place of Regular.

Optional `primary_color`, `accent_color` (hex colors), and `font_family` parameters let an agent pass selected values from a Figma design system into the formatter.

For PPTX input, the round trip preserves existing images, charts, and geometry while updating text. For PDF input, selectable text is rebuilt as an editable slide per PDF page; original layout and graphics are not preserved, and scanned PDFs need OCR first.

Figma MCP is a separate, optional MCP connection configured in `.vscode/mcp.json`. It can provide design context to the agent, but this server does not access Figma files or credentials directly. The agent can pass Figma-derived font and color tokens to `polish_deck`; it does not automatically convert Figma components into PowerPoint elements.

## Design Assumption: Large Type

Audiences stop paying attention when slide text is small, so every style option must keep type large. On a standard 10-inch-wide slide, use these floors:

- Titles: at least 40 pt
- Body text and bullets: at least 24 pt
- Captions, labels, sources, and slide numbers: at least 18 pt

When adopting an outside template, keep its palette and typefaces but scale its type up to these floors. For example, a 1920×1080 HTML template with 28 px body text is only about 10.5 pt on a slide.

## Existing Project Considered

[office-powerpoint-mcp-server](https://github.com/skittermonster/office-powerpoint-mcp-server) has useful theme-application and inspection capabilities. This project uses the same established `python-pptx` and MCP SDK foundations, while keeping the two-profile, folder-based workflow small and explicit instead of taking on that server's broader Office automation surface.
