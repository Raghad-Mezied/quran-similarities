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
  SectionType,
  ColumnBreak,
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
  return content.map((seg, i) => {
    const color = Object.keys(seg)[0];
    // Space separates segments, but the LAST segment gets none — otherwise a
    // trailing space sits between the last word and the closing ornate bracket ﴾.
    const sep = i === content.length - 1 ? "" : " ";
    return new TextRun({
      text: seg[color] + sep,
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

  // Pages ignores jc=center on import (it aligns purely by text direction), so a
  // centered title is not achievable from the file — right-aligned is the clean,
  // reliable RTL choice. The maroon underline is the paragraph's bottom border.
  const titlePara = new Paragraph({
    alignment: AlignmentType.RIGHT,
    bidirectional: true,
    spacing: { after: 200 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 12, color: "9C2A2A", space: 4 },
    },
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
        // Larger `before` puts a clear gap above each group's header, so the
        // groups of similarities read as visually separated blocks.
        spacing: { before: 320, after: 20 },
        children: [
          new TextRun({
            text: `(${mainText})`,
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
            new TextRun({ text: ":﴿", rightToLeft: true, font: "Arial" }),
            ...contentRuns(s.content, s.text_uthmani),
            new TextRun({ text: "﴾", rightToLeft: true, font: "Arial" }),
          ],
        }),
      );
    }
    return paras;
  }

  // The title spans the full page width; the verses below flow in two newspaper
  // columns with a vertical rule between them. So we use two sections: a
  // single-column one for the title, then a CONTINUOUS two-column one (continuous
  // keeps it on the same page rather than starting a new one) for the content.
  const titleSection = { children: [titlePara] };

  let contentSection;
  if (!blocks.length) {
    // No similarities — a single full-width message, no columns (matches the old
    // full-width box).
    contentSection = {
      properties: { type: SectionType.CONTINUOUS },
      children: [
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
      ],
    };
  } else {
    // Split the blocks into two balanced halves. Word fills the LEFT column
    // first (it lays newspaper columns out left-to-right regardless of the RTL
    // text), so to keep Arabic reading order — first half on the RIGHT — we put
    // the SECOND half before the column break (→ left column) and the FIRST half
    // after it (→ right column). A reader's eye starts on the right and reads the
    // first half first.
    const mid = Math.ceil(blocks.length / 2);
    const firstHalf = blocks.slice(0, mid).flatMap(blockParagraphs);
    const secondHalf = blocks.slice(mid).flatMap(blockParagraphs);

    const contentParas = [
      ...secondHalf,
      new Paragraph({ children: [new ColumnBreak()] }),
      ...firstHalf,
    ];

    contentSection = {
      properties: {
        type: SectionType.CONTINUOUS,
        // count: 2 → two equal columns; separate: true → draw the dividing line;
        // space → gap between the columns (in twips).
        column: { count: 2, space: 540, separate: true, equalWidth: true },
      },
      children: contentParas,
    };
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
    sections: [titleSection, contentSection],
  });

  return { title, blob: await Packer.toBlob(doc) };
}
