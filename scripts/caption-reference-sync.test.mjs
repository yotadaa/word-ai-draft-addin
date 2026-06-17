import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildAutoCaptionReferenceSyncPlan,
  extractCaptionNumber,
  findRenderedPageSplitParagraphs,
  findFallbackPageSplitParagraphs,
  getCaptionParagraphFormat
} from "../src/office.ts";

test("extractCaptionNumber reads numbered table and figure captions", () => {
  assert.equal(extractCaptionNumber("Tabel 3.10 Nilai crisp mutation rate", "Tabel"), "3.10");
  assert.equal(
    extractCaptionNumber("Gambar 3.13 Contoh penerapan AG berbasis website", "Gambar"),
    "3.13"
  );
});

test("extractCaptionNumber ignores plain title text", () => {
  assert.equal(extractCaptionNumber("Nilai crisp mutation rate", "Tabel"), "");
});

test("buildAutoCaptionReferenceSyncPlan maps old caption number to generated auto number", () => {
  assert.deepEqual(
    buildAutoCaptionReferenceSyncPlan({
      label: "Tabel",
      kind: "table",
      objectIndex: 7,
      oldCaptionText: "Tabel 3.10 Basis Aturan Fuzzy untuk Crossover Rate",
      newCaptionText: "Tabel 8. Basis Aturan Fuzzy untuk Crossover Rate"
    }),
    {
      label: "Tabel",
      kind: "table",
      oldNumber: "3.10",
      newNumber: "8",
      mentionText: "Tabel 3.10",
      bookmarkName: "auto_tbl_7_3_10"
    }
  );
});

test("buildAutoCaptionReferenceSyncPlan skips captions without old or generated numbers", () => {
  assert.equal(
    buildAutoCaptionReferenceSyncPlan({
      label: "Gambar",
      kind: "figure",
      objectIndex: 2,
      oldCaptionText: "Contoh penerapan AG berbasis website",
      newCaptionText: "Gambar 4. Contoh penerapan AG berbasis website"
    }),
    null
  );
});

test("getCaptionParagraphFormat centers converted captions", () => {
  assert.deepEqual(getCaptionParagraphFormat(), {
    fontName: "Bookman Old Style",
    fontSize: 10,
    italic: false,
    alignment: "centered"
  });
});

test("findRenderedPageSplitParagraphs detects unfinished sentence completed after page break", () => {
  const ooxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p><w:r><w:t>Kalimat sebelumnya sudah selesai.</w:t></w:r></w:p>
        <w:p>
          <w:r><w:t>Tabel 8 menjelaskan aturan fuzzy yang menghubungkan tingkat keragaman populasi dengan besarnya crossover rate yang harus</w:t></w:r>
          <w:r><w:lastRenderedPageBreak/></w:r>
          <w:r><w:t> diterapkan pada Algoritma Genetika.</w:t></w:r>
        </w:p>
      </w:body>
    </w:document>`;

  assert.deepEqual(findRenderedPageSplitParagraphs(ooxml), [
    {
      paragraphIndex: 1,
      beforeBreakText:
        "Tabel 8 menjelaskan aturan fuzzy yang menghubungkan tingkat keragaman populasi dengan besarnya crossover rate yang harus",
      afterBreakText: "diterapkan pada Algoritma Genetika.",
      splitSentence:
        "Tabel 8 menjelaskan aturan fuzzy yang menghubungkan tingkat keragaman populasi dengan besarnya crossover rate yang harus diterapkan pada Algoritma Genetika."
    }
  ]);
});

test("findRenderedPageSplitParagraphs ignores page break after completed sentence", () => {
  const ooxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r><w:t>Kalimat ini selesai.</w:t></w:r>
          <w:r><w:lastRenderedPageBreak/></w:r>
          <w:r><w:t> Kalimat baru ada di halaman berikutnya.</w:t></w:r>
        </w:p>
      </w:body>
    </w:document>`;

  assert.deepEqual(findRenderedPageSplitParagraphs(ooxml), []);
});

test("findFallbackPageSplitParagraphs returns long sentences when Word exposes no rendered break markers", () => {
  const paragraphs = [
    "Kalimat pendek sudah selesai.",
    "Tabel 8 menjelaskan aturan fuzzy yang menghubungkan tingkat keragaman populasi dengan besarnya crossover rate yang harus diterapkan pada Algoritma Genetika ketika diversity rendah sehingga dibutuhkan crossover tinggi untuk meningkatkan eksplorasi dan mencegah solusi lokal.",
    "Tabel 8. Basis Aturan Fuzzy untuk Crossover Rate"
  ];

  assert.deepEqual(findFallbackPageSplitParagraphs(paragraphs), [
    {
      paragraphIndex: 1,
      splitSentence:
        "Tabel 8 menjelaskan aturan fuzzy yang menghubungkan tingkat keragaman populasi dengan besarnya crossover rate yang harus diterapkan pada Algoritma Genetika ketika diversity rendah sehingga dibutuhkan crossover tinggi untuk meningkatkan eksplorasi dan mencegah solusi lokal."
    }
  ]);
});
