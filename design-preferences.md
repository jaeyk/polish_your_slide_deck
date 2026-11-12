# Design Preferences

Style rules for decks polished in this project. Every style option, template, or `polish_deck` change should follow them.

**Default style: Newsroom Light** (`polish_deck`'s default `newsroom_light` profile). It uses a cool white background (#F5F7FA), near-black text (#141A22), and Carolina Navy (#13294B) emphasis blocks with light text. The user chose it on 2026-11-11, partly because white screenshots and charts blend into it. Civic Signal (`civic_signal`: navy, white, signal yellow) is the dark alternative. Font: **Overpass**, based on Highway Gothic, the US road-sign typeface, chosen for its public-infrastructure voice. Libre Franklin was rejected as too "Claude"-looking.

## Academic, authoritative, never Beamer

The audience is academic, so slides should look authoritative.

- **Use** restrained editorial styling: heavy sans-serif type (serif titles were tried and read as generic), one authoritative accent color, generous whitespace, and an assertion–evidence structure (one claim per slide in a full-sentence headline, with the evidence below it).
- **Avoid Beamer tropes:** title bars, navigation dots, boxed theorem-style blocks, author/title/date footers, and the Computer Modern look.
- **Avoid loud, playful palettes.** The neo-brutalist "creative-mode" template (cream background, four saturated accents, Archivo Black) was judged "okay but not so great" for academic audiences.

## Not generic, not "Claude-looking"

The style should not look like a default AI-generated deck. Avoid this combination:

- warm cream or paper backgrounds
- serif display titles with italic emphasis
- warm gold or clay accents
- small uppercase labels above a short rule
- thin hairline-divided rows

Favored direction: a newsroom or data-journalism voice, with heavy sans-serif type, one cool accent color, solid inverted highlight blocks, and the institution's own colors (Carolina Navy #13294B, Carolina Blue #7BAFD4). A dark lecture-hall background and the light Carolina Navy palette are both liked.

## Balanced vertical spacing

Don't leave a large empty band at the top with content sitting low. Center each slide's content block vertically, placed slightly above true center (about 12 px higher on a 540 px-tall slide) so the top and bottom margins look even. Slide numbers don't count as content.

## No faint text

Text that is too faint is as bad as text that is too small.

- **Contrast:** every text color needs at least 7:1 contrast against its background (WCAG AAA). Muted mid-greys such as #5B6470 (5.6:1 on paper) are too faint; use a dark slate such as #2F3742 (11.2:1) for secondary text.
- **Weight:** never use Light or Thin weights. Sans-serif text should be SemiBold or heavier, because thin strokes wash out on projectors.
- **Pale tints** are only for highlight backgrounds, never for text the audience must read.

## No redundant decoration

Don't add ornaments that repeat what the slide already says, such as an oversized background numeral next to "Study 1". Every element should carry information.

## Large type

Audiences stop paying attention when slide text is small. The size floors are in [README.md](README.md#design-assumption-large-type): titles at least 40 pt, body text at least 24 pt, and labels at least 18 pt.
