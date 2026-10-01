/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { TableLayout } from './types.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

function layoutTable(tableXml: string, mode = 14): TableLayout {
  const parts = new Map([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    </Types>`],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    </Relationships>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${tableXml}
      <w:sectPr><w:pgSz w:w="12240" w:h="14400"/>
        <w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/>
      </w:sectPr></w:body></w:document>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:style w:type="table" w:styleId="LogicalTable"><w:name w:val="Logical table"/></w:style></w:styles>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:compat>
      <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>
    </w:compat></w:settings>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map(
    [...parts].map(([name, xml]) => [name, encoder.encode(xml)]),
  )));
  try {
    const model = normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
    const context = {
      font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
      measureText: (text: string) => ({
        width: [...text].length * 6,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      }),
    } as unknown as CanvasRenderingContext2D;
    const result = layoutDocument(model, createLayoutServices(model, { measureContext: context }), {
      currentDateMs: 0,
    });
    const table = result.pages[0]?.layers.body.find((block) => block.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected retained parsed table');
    return table;
  } finally {
    archive.free();
  }
}

const margins = (left = 360) => `<w:tblCellMar><w:left w:type="dxa" w:w="${left}"/><w:right w:type="dxa" w:w="180"/></w:tblCellMar>`;
const cell = (width: number, left?: number) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${width}"/>${left === undefined ? '' : `<w:tcMar><w:left w:type="dxa" w:w="${left}"/></w:tcMar>`}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`;
function table(rows: string, options = ''): string {
  // ECMA-376 CT_TblPrBase orders style/width/placement before layout/margins.
  const propertyOrder = ['tblStyle', 'bidiVisual', 'tblW', 'jc', 'tblCellSpacing', 'tblInd'];
  const orderedOptions = [...options.matchAll(/<w:(\w+)\b[^>]*\/>/g)]
    .sort((a, b) => propertyOrder.indexOf(a[1] ?? '') - propertyOrder.indexOf(b[1] ?? ''))
    .map(([xml]) => xml).join('');
  return `<w:tbl><w:tblPr>${orderedOptions}<w:tblLayout w:type="fixed"/>${margins()}</w:tblPr><w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="3600"/></w:tblGrid>${rows}</w:tbl>`;
}
const indent = '<w:tblInd w:type="dxa" w:w="108"/>';
const row = `<w:tr>${cell(2160)}${cell(3600)}</w:tr>`;
beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('Word table origin and occupied grid', () => {
  it.each([11, 12, 14, 15])('mode %s distinguishes missing indentation and authored leading indentation', (mode) => {
    const missing = layoutTable(table(row), mode);
    const authored = layoutTable(table(row, indent), mode);
    expect(missing.flowBounds.xPt).toBe(72);
    expect(authored.flowBounds.xPt).toBeCloseTo(mode === 15 ? 77.4 : 59.4, 8);
    expect(layoutTable(table(row, indent.replace('108', '0')), mode).flowBounds.xPt)
      .toBeCloseTo(mode === 15 ? 72 : 54, 8);
    const rtl = layoutTable(table(row, indent + '<w:bidiVisual/>'), mode);
    expect(rtl.flowBounds.xPt).toBeCloseTo(mode === 15 ? 246.6 : 264.6, 8);
  });
  it.each(['center', 'right'])('%s ignores indentation and uses the first cell margin only at the logical end', (jc) => {
    const rows = `<w:tr>${cell(2160, 540)}${cell(3600)}</w:tr><w:tr>${cell(2160, 120)}${cell(3600)}</w:tr>`;
    const options = indent + `<w:jc w:val="${jc}"/>`;
    const legacy = layoutTable(table(rows, options));
    expect(legacy.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual(jc === 'center' ? [162, 162] : [279, 258]);
    const rtl = layoutTable(table(rows, options + '<w:bidiVisual/>'));
    expect(rtl.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual(jc === 'center' ? [342, 342] : [225, 246]);
    const modern = layoutTable(table(rows, options), 15);
    expect(modern.rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual(jc === 'center' ? [162, 162] : [252, 252]);
  });
  it('first-row exceptions anchor every leading row with the first cell override', () => {
    const rows = `<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/>${margins(240)}</w:tblPrEx>${cell(2160, 540)}${cell(3600)}</w:tr>${row}`;
    expect(layoutTable(table(rows, indent)).rows.map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([81, 81]);
  });
  it('preserves the previous origin for unresolved missing-margin and spaced classes', () => {
    const missing = table(row, indent + '<w:jc w:val="right"/>').replace(margins(), '');
    expect(layoutTable(missing).flowBounds.xPt).toBeCloseTo(257.4, 8);
    const logicalOverride = table(`<w:tr>${cell(2160, 540).replace('<w:left ', '<w:start ')}${cell(3600)}</w:tr>`, indent);
    expect(layoutTable(logicalOverride).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const unpreferred = table(row.replaceAll('w:type="dxa"', 'w:type="auto"'), indent);
    expect(layoutTable(unpreferred).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const autoFit = table(row, indent).replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"');
    expect(layoutTable(autoFit).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const cleared = table(`<w:tr><w:tblPrEx><w:tblW w:type="dxa" w:w="0"/></w:tblPrEx>${cell(2160)}${cell(3600)}</w:tr>`, indent + '<w:tblW w:type="dxa" w:w="5760"/>')
      .replace('w:tblLayout w:type="fixed"', 'w:tblLayout w:type="autofit"');
    expect(layoutTable(cleared).flowBounds.xPt).toBeCloseTo(77.4, 8);
    const spaced = table(row, indent + '<w:tblCellSpacing w:type="dxa" w:w="120"/>');
    expect(layoutTable(spaced).flowBounds.xPt).toBeCloseTo(77.4, 8);
  });
  it('preserves member indentation outside the measured logical-table scope', () => {
    const style = '<w:tblStyle w:val="LogicalTable"/>';
    const exception = row.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>');
    expect(layoutTable(table(exception, style) + table(row, style), 13).rows
      .map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([108, 72]);
    const unpreferred = row.replaceAll('w:type="dxa"', 'w:type="auto"');
    expect(layoutTable(table(exception, style) + table(unpreferred, style), 14).rows
      .map((r) => r.cells[0]?.flowBounds.xPt)).toEqual([108, 72]);
  });
  it('preserves the occupied-grid solver for unmeasured percentage, negative-width and nested classes', () => {
    const skipped = `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const gridTable = table(skipped, '<w:tblW w:type="auto" w:w="0"/>')
      .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    const percentage = gridTable.replace('w:type="auto" w:w="0"', 'w:type="pct" w:w="5000"');
    expect(layoutTable(percentage).columnWidthsPt).toEqual(layoutTable(percentage, 13).columnWidthsPt);
    const negative = gridTable.replace('<w:tcW w:type="dxa" w:w="2160"', '<w:tcW w:type="dxa" w:w="-2160"');
    expect(layoutTable(negative).columnWidthsPt).toEqual(layoutTable(negative, 13).columnWidthsPt);
    const outer = table(`<w:tr>${cell(9360).replace('<w:p>', gridTable + '<w:p>')}</w:tr>`)
      .replace('<w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="3600"/>', '<w:tblGrid><w:gridCol w:w="9360"/>');
    const nested = layoutTable(outer).rows[0]?.cells[0]?.blocks.find((block) => block.layout.kind === 'table')?.layout;
    expect(nested?.kind === 'table' ? nested.columnWidthsPt : null).toEqual([36, 108, 180]);
  });
  it.each([14, 15])('mode %s drops a universally skipped leading fixed track, while an occupied preceding row preserves it', (mode) => {
    const skipped = (before: number) => `<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:type="dxa" w:w="${before}"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const full = `<w:tr>${cell(720)}${cell(2160)}${cell(3600)}</w:tr>`;
    const gridTable = (rows: string, preferred: boolean) => table(rows, '<w:tblInd w:type="dxa" w:w="0"/>' + (preferred ? '<w:tblW w:type="dxa" w:w="6480"/>' : '<w:tblW w:type="auto" w:w="0"/>'))
      .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    for (const before of [0, 360, 720]) {
      const automatic = layoutTable(gridTable(skipped(before), false), mode);
      expect(automatic.columnWidthsPt).toEqual([0, 108, 180]);
      const preferred = layoutTable(gridTable(skipped(before), true), mode);
      expect(preferred.columnWidthsPt).toEqual([0, 121.5, 202.5]);
      const occupied = layoutTable(gridTable(full + skipped(before), true), mode);
      expect(occupied.columnWidthsPt).toEqual([36, 108, 180]);
      expect(occupied.rows[1]?.cells[0]?.flowBounds.xPt).toBeCloseTo((mode === 15 ? 72 : 54) + 36, 8);
    }
  });
  it.each([11, 12, 14, 15])('mode %s lays out split logical tables like a single tbl', (mode) => {
    const skipped = `<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:type="dxa" w:w="720"/></w:trPr>${cell(2160)}${cell(3600)}</w:tr>`;
    const full = `<w:tr>${cell(720)}${cell(2160)}${cell(3600)}</w:tr>`;
    const override = full.replace(cell(720), cell(720, 540));
    const exception = full.replace('<w:tr>', '<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="720"/></w:tblPrEx>');
    const unpreferred = skipped.replace('<w:tcW w:type="dxa" w:w="2160"', '<w:tcW w:type="auto" w:w="0"');
    const gridTable = (rows: string, preferred: boolean) => table(rows,
      '<w:tblStyle w:val="LogicalTable"/><w:tblInd w:type="dxa" w:w="0"/>'
      + (preferred ? '<w:tblW w:type="dxa" w:w="6480"/>' : '<w:tblW w:type="auto" w:w="0"/>'))
      .replace('<w:tblGrid>', '<w:tblGrid><w:gridCol w:w="720"/>');
    const geometry = (layout: TableLayout) => ({
      columns: layout.columnWidthsPt,
      bounds: layout.flowBounds,
      rows: layout.rows.map((r) => r.cells.map((c) => c.flowBounds)),
    });
    for (const preferred of [false, true]) {
      // Occupancy on either side of an authored seam must preserve track zero;
      // all-skipped members must still normalize. Different first-cell margins
      // also exercise the logical first-row leading anchor.
      for (const rows of [[full, skipped], [skipped, full], [skipped, skipped], [override, full], [exception, full], [unpreferred, skipped]]) {
        const single = layoutTable(gridTable(rows.join(''), preferred), mode);
        const split = layoutTable(rows.map((r) => gridTable(r, preferred)).join(''), mode);
        expect(geometry(split)).toEqual(geometry(single));
        if (rows[0] === full && rows[1] === skipped) {
          expect(split.rows[1]?.cells.map((c) => c.flowBounds.xPt))
            .toEqual(mode === 15 ? [108, 216] : [90, 198]);
          const keepNext = '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>lead</w:t></w:r></w:p>';
          expect(geometry(layoutTable(keepNext + rows.map((r) => gridTable(r, preferred)).join(''), mode)))
            .toEqual(geometry(layoutTable(keepNext + gridTable(rows.join(''), preferred), mode)));
        }
      }
    }
  });
});
