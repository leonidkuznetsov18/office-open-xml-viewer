# Positioned text frames in Word table cells

## Scope and evidence

ECMA-376 §17.3.1.11 defines `w:framePr` on a paragraph, including its
horizontal and vertical anchors, offsets, size, and wrapping. A paragraph
inside `w:tc` can carry the same property. The DOCX parser already preserves
it in `DocParagraph.framePr`; the missing capability is in table layout.

Word-exported OOXML controls show a narrow placement rule. In the first cell,
an initial `framePr` paragraph with positive `w` moves to its page or text
anchor; adjacent paragraphs with the same frame property form one group.
Absent or zero `w` remains in ordinary cell flow. A positive-width group after
an ordinary paragraph, or in a neighboring cell, also stays in cell flow.
The controls varied width at zero and one twip, paragraph order, and cell
index. The moved case cannot be represented by stacking all cell blocks.

In split-row page-anchor and text-anchor controls, the later framed group
stays in flow and appears immediately after the repeated header on the
continuation page. Its page/text anchor offsets do not relocate it there.
The renderer must preserve ordinary row splitting and repeated-header order
for that case. These results replace the earlier assumption that a split-page
frame is independently positioned against the destination page.

## Current boundary

`layout/table-acquisition.ts` acquires each cell content element independently
and sends the resulting paragraph or nested-table layout to `cellBlocks`.
`layout/table.ts::resolveCellFlow` assigns every paragraph an `offsetPt` and
`advancePt` in the cell's vertical stack. `TableCellBlockLayout` stores only
that flow placement. Neither acquisition nor the retained table/cell layout
has an out-of-flow frame group, frame anchor, exclusion, or page-fragment
owner. Body frames use `acquireRetainedFrameGroup` and a body float registry;
applying that body coordinate system to a cell would place text outside its
owner and would not survive table fragmentation correctly.

## Required generic capability

1. Group adjacent cell paragraphs with the same effective `CT_FramePr` value,
   using the existing frame identity rule. Create a separate placement only
   for a positive-width group at the start of the first cell; keep the measured
   counterexamples in flow. Retain one owner and the member fragments.
2. Acquire the group's contents at its frame width, and retain its positioned
   bounds, exclusion bounds, anchor identity, wrap mode, and source order as a
   `CellFramePlacement` in the table layout. Keep it separate from the normal
   `TableCellBlockLayout` advance. The normal cell paragraphs still determine
   the cell's flow height.
3. Resolve text-relative anchors after the cell's final position is known;
   resolve page- and margin-relative anchors against the destination page of
   the table fragment. Register exclusions in the correct cell/fragment flow
   domain before laying out text that wraps around the frame. Do not let a
   cell-local exclusion silently change a neighboring cell.
4. Carry frame placement through row splitting, repeated headers, vertical
   cell text transforms, cell clipping, and paint order. Invalidate or
   reacquire page-dependent placement when a table fragment moves to another
   page; the current reusable table acquisition cannot cache it unchanged.
5. Keep the parser/model path shared by OOXML and optional binary readers.
   Binary-specific placement logic belongs only in the legacy acquisition
   layer after the generic table capability exists.

## Acceptance

Add focused OOXML tests for absent and zero-width frames, first-cell initial
page/text anchors, adjacent group members, following cell text, preceding
ordinary text, a neighboring cell, and a split row. Compare those controls
with Word PDFs, including the continuation immediately after a repeated
header.
Run a private DOCX self-VRT against the exact `origin/main` baseline, then
adjudicate every changed page against Word. Only after that can a binary DOC
cell-frame fact be projected without silently reverting to cell flow.

The implementation follows only the measured placement envelope above.
