# Positioned text frames in Word table cells

## Scope and evidence

ECMA-376 §17.3.1.11 defines `w:framePr` on a paragraph, including its
horizontal and vertical anchors, offsets, size, and wrapping. A paragraph
inside `w:tc` can carry the same property. The DOCX parser already preserves
it in `DocParagraph.framePr`; the missing capability is in table layout.

Word-exported PDFs of three OOXML controls hold the table, text, and frame
dimensions constant while changing only the cell paragraph's `framePr`:
absent, `hAnchor/vAnchor=page`, and `hAnchor/vAnchor=text`. The framed text is
ordinary cell flow without the property, moves to the page-origin offset for
the page anchor, and moves relative to its text anchor for the text anchor.
The following cell paragraph moves with the frame. Those positions cannot be
represented by treating both paragraphs as stacked cell blocks. A separate
control with a frame in a table spanning page boundaries is still needed to
settle which fragment owns each frame and wrap exclusion.

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
   using the existing frame identity rule. Retain one owner and the member
   fragments in the cell acquisition graph.
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

Add focused OOXML tests for absent, page-anchored, and text-anchored frames in
a cell, adjacent frame paragraphs, wrapping around following cell text, a
neighboring cell, and a split row. Compare those controls with Word PDFs.
Run a private DOCX self-VRT against the exact `origin/main` baseline, then
adjudicate every changed page against Word. Only after that can a binary DOC
cell-frame fact be projected without silently reverting to cell flow.

This branch records the missing capability and its design; it does not claim
cell-frame rendering support.
