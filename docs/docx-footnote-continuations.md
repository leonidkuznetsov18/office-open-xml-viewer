# Page-bottom footnote continuation

The DOCX layout engine can continue a footnote onto the next physical page when
`allowFootnoteContinuation: true` is passed to `DocxDocument.load`. The default
remains `false` while the renderer lacks row-aware splitting of footnote tables
and support for authored continuation-separator stories. A caller should enable
this option only for documents whose notes contain paragraphs and whose page
geometry is stable across the affected page boundary.

ECMA-376 Part 1 §17.11.1 describes a footnote that continues onto the next
page. §17.11.21 places notes at the page bottom. The continuation keeps the
reference on its original body page, partitions the note at a complete line,
and resumes the remaining lines in a new note band on the following physical
page. The continuation page uses the full text width for its separator, as in
Word's default continuation separator. An ordinary first-page separator uses
one third of the text width.

The controls are Word-created documents and their PDFs. They establish that a
note can continue even when its body reference remains on the prior page, and
that reducing the note's font size can make the continuation disappear. A
counterexample with a short note retains it as a whole on the next page. The
exact body-text break and note admission remain sensitive to line measurement;
this option does not claim to match every Word pagination decision.

The private DOCX self regression against the pre-feature renderer is exact
with the option at its default value. Focused layout tests cover line ownership,
empty trailing paragraphs, reference relocation when the first line cannot fit,
and the public paginator path.
