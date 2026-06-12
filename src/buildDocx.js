// Build a real .docx for one surah — entirely in the browser.
// This is the port of the old server.js `/word/:num` route: the `docx` library
// runs client-side, so `Packer.toBlob()` gives us a downloadable file with no backend.
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  BorderStyle,
  Table,
  TableRow,
  TableCell,
  TableLayoutType,
  WidthType,
  VerticalAlign,
} from "docx";
import { SURAHS, colorHex } from "./surah-data.js";

// ---- Build coloured runs from a verse's `content` segments ----
function contentRuns(content, fallback, extra = {}) {
  if (!Array.isArray(content) || !content.length) {
    return [
      new TextRun({
        text: fallback || "",
        rightToLeft: true,
        font: "Arial",
        ...extra,
      }),
    ];
  }
  return content.map((seg) => {
    const color = Object.keys(seg)[0];
    return new TextRun({
      text: seg[color] + " ",
      color: colorHex(color),
      rightToLeft: true,
      font: "Arial",
      ...extra,
    });
  });
}

// ---- Extract the verses that HAVE similarities (skip the rest) ----
function buildBlocks(data) {
  const blocks = [];
  for (const verse of data || []) {
    const els = verse.elements || [];
    const main = els.find((e) => !e.sim_chapter_name) || els[0];
    const sims = els.filter((e) => e.sim_chapter_name);
    if (!main || !sims.length) continue;
    blocks.push({ main, sims });
  }
  return blocks;
}

// ---- Generate a real .docx for one surah, returning a Blob ----
export async function buildDocx(num, data) {
  const name = SURAHS[num - 1] || `سورة ${num}`;
  const title = `متشابهات سورة ${name}`;
  const blocks = buildBlocks(data);

  // Shared table geometry. Pages only right-aligns text inside an RTL table —
  // NOT in a standalone paragraph — so even the title goes in a table cell.
  const COL_W = 4500; // ~half the usable page width (works for Letter & A4)
  const TABLE_W = COL_W * 2;
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  // A full-width single-cell RTL table — used to force right-alignment on
  // content that isn't part of the 2-column grid (title, messages).
  const fullWidthBox = (paras, bottomBorder) =>
    new Table({
      visuallyRightToLeft: true,
      alignment: AlignmentType.RIGHT,
      layout: TableLayoutType.FIXED,
      columnWidths: [TABLE_W],
      width: { size: TABLE_W, type: WidthType.DXA },
      borders: {
        top: noBorder,
        left: noBorder,
        right: noBorder,
        insideHorizontal: noBorder,
        insideVertical: noBorder,
        bottom: bottomBorder || noBorder,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: TABLE_W, type: WidthType.DXA },
              margins: { top: 40, bottom: 100, left: 140, right: 140 },
              children: paras,
            }),
          ],
        }),
      ],
    });

  // Pages ignores jc=center on import (it aligns purely by text direction), so a
  // centered title is not achievable from the file — right-aligned is the clean,
  // reliable RTL choice. The maroon underline is the box's bottom border.
  const titlePara = new Paragraph({
    alignment: AlignmentType.RIGHT,
    bidirectional: true,
    children: [
      new TextRun({
        text: title,
        bold: true,
        size: 36,
        color: "1F3864",
        rightToLeft: true,
        font: "Arial",
      }),
    ],
  });

  const titleBox = fullWidthBox([titlePara], {
    style: BorderStyle.SINGLE,
    size: 12,
    color: "9C2A2A",
  });

  // Turn one block (a main verse + its similar verses) into right-aligned paragraphs.
  function blockParagraphs(b) {
    const paras = [];
    // main verse — bold maroon group header, wrapped in parentheses
    const mainText =
      Array.isArray(b.main.content) && b.main.content.length
        ? b.main.content.map((seg) => seg[Object.keys(seg)[0]]).join(" ")
        : b.main.text_uthmani || "";
    paras.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        bidirectional: true,
        spacing: { before: 120, after: 20 },
        children: [
          new TextRun({
            text: `( ${mainText} )`,
            bold: true,
            color: "9C2A2A",
            size: 28,
            rightToLeft: true,
            font: "Arial",
          }),
        ],
      }),
    );
    if (b.main.comment) {
      paras.push(
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          bidirectional: true,
          children: [
            new TextRun({
              text: b.main.comment,
              italics: true,
              color: "777777",
              rightToLeft: true,
              font: "Arial",
            }),
          ],
        }),
      );
    }
    // similar verses — manual bullet (reliable in RTL),  format:  • surah aya : ( aya )
    for (const s of b.sims) {
      paras.push(
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          bidirectional: true,
          spacing: { before: 20, after: 60 },
          children: [
            // Bullet glued onto the Arabic ref so it sits inside a strong RTL run
            // and resolves to the RIGHT (a standalone "•" run is bidi-neutral and
            // Pages drops it on the left).
            new TextRun({
              text: `‏• ${s.sim_chapter_name} ${s.sim_verse_number}‏`,
              bold: true,
              color: "C00000",
              rightToLeft: true,
              font: "Arial",
            }),
            new TextRun({ text: " : ﴿", rightToLeft: true, font: "Arial" }),
            ...contentRuns(s.content, s.text_uthmani),
            new TextRun({ text: "﴾", rightToLeft: true, font: "Arial" }),
          ],
        }),
      );
    }
    return paras;
  }

  const children = [titleBox];

  if (!blocks.length) {
    children.push(
      fullWidthBox([
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          bidirectional: true,
          children: [
            new TextRun({
              text: "لا توجد آيات متشابهة في هذه السورة.",
              rightToLeft: true,
              font: "Arial",
              color: "999999",
            }),
          ],
        }),
      ]),
    );
  } else {
    // Split the blocks into two balanced halves — a real 2-column table (like the
    // web preview). First half reads first, so it goes in the RIGHT column.
    const mid = Math.ceil(blocks.length / 2);
    const rightParas = blocks.slice(0, mid).flatMap(blockParagraphs);
    const leftParas = blocks.slice(mid).flatMap(blockParagraphs);
    if (!leftParas.length) leftParas.push(new Paragraph({ text: "" }));

    const cell = (paras) =>
      new TableCell({
        width: { size: COL_W, type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        margins: {
          top: 80,
          bottom: 40,
          left: 250,
          right: 250,
        },
        children: paras,
      });

    // RTL table (bidiVisual). This is the Word/Pages standard for right-to-left
    // tables: cells lay out right-to-left (first cell = right column) AND cell
    // content is right-aligned. Without it, Pages treats the table as LTR and
    // left-aligns everything.
    children.push(
      new Table({
        visuallyRightToLeft: true,
        alignment: AlignmentType.RIGHT,
        layout: TableLayoutType.FIXED,
        columnWidths: [COL_W, COL_W],
        width: { size: COL_W * 2, type: WidthType.DXA },
        borders: {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: {
            style: BorderStyle.SINGLE,
            size: 2,
            color: "C08080",
          },
        },
        rows: [new TableRow({ children: [cell(rightParas), cell(leftParas)] })],
      }),
    );
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Arial",
            size: 26,
            rightToLeft: true,
            language: { value: "ar-SA" },
          },
          paragraph: {
            bidirectional: true,
            alignment: AlignmentType.RIGHT,
          },
        },
      },
    },
    sections: [{ children }],
  });

  return { title, blob: await Packer.toBlob(doc) };
}
