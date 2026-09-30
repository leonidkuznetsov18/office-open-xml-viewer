# Tab fitting differential check

This separate integration gate compares the current DOCX layout with a clean
checkout at the current `origin/main`, using identical parsed facts and text
metrics. Install dependencies in both checkouts and build the candidate WASM
before running:

```sh
DOCX_TAB_BASELINE_CHECKOUT=<baseline-checkout> pnpm exec vitest run \
  --config packages/docx/tests/differential/tab-fitting.config.ts
```

It covers every no-float case from the float/tab matrix and generated ordinary
RTL/LTR lines with 1–3 tabs, all tab alignments, custom stops, prefixes, and
Latin/CJK/Thai content. Line partitions and placement geometry must agree unless
main allocates outside the paragraph band or the case is one of the five short
Word-backed RTL positional exceptions. Every candidate must preserve the text
and stay in band. An optional `DOCX_TAB_DIFFERENTIAL_REPORT` path saves counts
outside the checkout. This gate runs separately from `pnpm test` because it
requires the second checkout.
