# PPTX empty East Asian font slot compatibility

A run whose East Asian (`a:ea`) font slot is empty now draws its East
Asian-slot characters the way PowerPoint does. Fonts missing from the reference
catalogue no longer switch a whole text body to the ordinary line model. No
migration is required.

## Specification boundary

ECMA-376 Part 1 §21.1.2.3 assigns characters to the Latin, East Asian, complex
script and symbol slots. §21.1.2.3.1/.3, §20.1.4.1.16–.18/.24/.25 and
§20.1.10.81 define the slot fonts, theme tokens and empty typefaces.
§21.1.2.2.5/.11 make line size and spacing line-local.

Two things below are **observed PowerPoint compatibility**, not normative
rules:

- which face draws a glyph when the East Asian slot is empty;
- where the baseline splits inside a line.

## Office evidence

PowerPoint for Mac 16.113 opened synthetic control decks and exported tagged
PDFs with "best for electronic distribution". The controls held each factor
fixed while varying one at a time:

- 64 control/test pairs over four empty-slot decks;
- 29 fallback-scope pairs.

Faces were read from the embedded fonts' name tables, not from PDF font labels.
Baselines were read from the text state. The decks and PDFs are local
verification artifacts, not redistributable baselines.

- **Selected face.** With an empty, omitted or theme-empty `a:ea`, the selected
  face S is:
  - the run's complex-script (`a:cs`) face, when one is named (directly or
    through a theme token);
  - otherwise the run's Latin face.

  S draws every East Asian-slot glyph it covers, including Han and kana when it
  maps them. Language (en/ja/zh/ko), the theme Latin face and the Latin face's
  own coverage play no part when a cs face is named.
- **Symbol fallback.** A non-CJK glyph that S lacks is drawn in Calibri (§),
  even when no Calibri appears in the deck, or in Cambria Math (◆ ■).
- **CJK fallback.** For a CJK glyph that S lacks:
  - A Far-East face that maps basic CJK uses its own script chain, whichever
    slot it fills. A Japanese face sends simplified Han to Microsoft JhengHei
    and Hangul to Malgun Gothic.
  - A face that maps no CJK Unified Ideograph falls back by its PANOSE serif
    style. Styles 1–10 go to MS Mincho; 11–15 and unknown faces go to MS Gothic.
  - Whether a face maps basic CJK comes from its installed cmap, recorded in the
    reference catalogue.
- **Synthetic styles.** A family without an italic face is drawn upright with a
  synthetic slant. Its vertical metrics are those of the upright resource.
- **Line metrics.** A line is sized by the OS/2 metrics of the resources that
  draw it: usWinAscent / usWinDescent, or the typo metrics plus line gap under
  USE_TYPO_METRICS. A face changes only its own line; earlier and later lines
  keep their baselines.

## Owner decisions (accepted platform differences)

- **(B) Second CJK fallback.** After the first fallback of a Far-East face that
  maps no basic CJK, the next face depends on the font itself. SimSun-ExtB and
  MingLiU-ExtB differ, and no general font property explains the difference.
  It is left to the browser's own glyph fallback. No face name is encoded.
- **(c) Unknown drawing face.** The browser does not report which installed
  face draws a fallback glyph.
  - Such a glyph adds no face to its line, and the line keeps the metric model
    of its known faces. Only a line with no known face at all uses the ordinary
    model.
  - Synthetic italic uses the browser's own oblique, not PowerPoint's 0.3333
    shear.
  - The catalogue does not record symbol coverage. A symbol that S lacks is
    therefore still sized by S, while the browser draws it with the symbol
    fallback where those fonts are installed.

## Implementation

`packages/pptx/src/east-asian-default.ts` builds the stack in this order:
the selected face, then the symbol fallback, then the CJK fallback faces. It
also names the drawing face when the renderer can know it.

`renderer.ts` uses that stack for measuring, wrapping, painting and stacked or
vertical text. It sizes each line from its known faces only.

Deck-embedded fonts are sized from their own font parts' OS/2 tables. These
are parsed when the font is registered, in the window and in the render worker.

Structural cases outside the measured controls still use the ordinary model
for the whole body:

- equations;
- markers taller than the text;
- `compatLnSpc="0"` without Excel tables;
- unresolved `fontAlgn` offsets.
