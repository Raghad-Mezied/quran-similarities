import { useState, useCallback } from "react";
import { SURAHS } from "./surah-data.js";
import { buildDocx } from "./buildDocx.js";

// Turn one surah's JSON into structured blocks — only verses that HAVE similarities.
// Mirrors buildBlocks in buildDocx.js so the on-screen preview matches the Word file.
function buildBlocks(data) {
  const blocks = [];
  for (const verse of data || []) {
    const els = verse.elements || [];
    const main = els.find((e) => !e.sim_chapter_name) || els[0];
    const sims = els.filter((e) => e.sim_chapter_name);
    if (!main || !sims.length) continue;

    const mainText =
      Array.isArray(main.content) && main.content.length
        ? main.content.map((seg) => seg[Object.keys(seg)[0]]).join(" ")
        : main.text_uthmani || "";
    blocks.push({ mainText, comment: main.comment, sims });
  }
  return blocks;
}

// Render the coloured segments of a verse (content[] = [{red:"…"},{black:"…"}]).
function Segments({ content, fallback }) {
  if (!Array.isArray(content) || !content.length) return <>{fallback || ""}</>;
  return content.map((seg, i) => {
    const color = Object.keys(seg)[0];
    // 'unknown' marks the "…" elision — render it grey, like the colour map does.
    const css = color === "unknown" ? "#888" : color;
    return (
      <span key={i} style={{ color: css }}>
        {seg[color]}{" "}
      </span>
    );
  });
}

function Block({ block }) {
  return (
    <div className="block">
      <p className="main">( {block.mainText} )</p>
      {block.comment && <p className="comment">{block.comment}</p>}
      <ul>
        {block.sims.map((s, i) => (
          <li key={i}>
            <span className="ref">
              {s.sim_chapter_name} {s.sim_verse_number}
            </span>{" "}
            : ( <Segments content={s.content} fallback={s.text_uthmani} /> )
          </li>
        ))}
      </ul>
    </div>
  );
}

function Preview({ blocks }) {
  if (!blocks.length) {
    return <div className="empty">لا توجد آيات متشابهة في هذه السورة.</div>;
  }
  // split in half — right column gets the first half (RTL reading order)
  const mid = Math.ceil(blocks.length / 2);
  const right = blocks.slice(0, mid);
  const left = blocks.slice(mid);
  return (
    <table className="cols" dir="rtl">
      <tbody>
        <tr>
          <td className="col">
            {right.map((b, i) => (
              <Block key={i} block={b} />
            ))}
          </td>
          <td className="divider"></td>
          <td className="col">
            {left.map((b, i) => (
              <Block key={i} block={b} />
            ))}
          </td>
        </tr>
      </tbody>
    </table>
  );
}

// Build the .docx in the browser and trigger a download (replaces the old /word/:num route).
async function downloadWord(num, data) {
  const { title, blob } = await buildDocx(num, data);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${title}.docx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export default function App() {
  const [active, setActive] = useState(null); // selected surah number
  const [current, setCurrent] = useState(null); // { num, title, blocks, data }
  const [status, setStatus] = useState("idle"); // idle | loading | error
  const [errorMsg, setErrorMsg] = useState("");

  const selectSurah = useCallback(async (num) => {
    setActive(num);
    setStatus("loading");
    setCurrent(null);
    try {
      const res = await fetch(`/data/${num}.json`);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();

      const blocks = buildBlocks(data);
      const title = `متشابهات سورة ${SURAHS[num - 1] || "سورة " + num}`;
      setCurrent({ num, title, blocks, data });
      setStatus("idle");

      // The original behaviour: pressing the surah button also generates the Word file.
      await downloadWord(num, data);
    } catch (err) {
      setErrorMsg(err.message);
      setStatus("error");
    }
  }, []);

  const docTitle =
    status === "loading"
      ? "…جارٍ التحميل"
      : status === "error"
        ? "خطأ"
        : current
          ? current.title
          : "اختر سورة من الأعلى…";

  return (
    <>
      <header>
        <h1>المتشابهات اللفظية في القرآن الكريم</h1>
        <p>اختر السورة لتوليد ملف Word بمتشابهاتها — ١ إلى ١١٤</p>
      </header>

      <div className="wrap">
        <div className="grid">
          {SURAHS.map((name, i) => {
            const num = i + 1;
            return (
              <button
                key={num}
                className={"surah-btn" + (active === num ? " active" : "")}
                onClick={() => selectSurah(num)}
              >
                <span className="num">{num}</span>
                <span className="name">{name}</span>
              </button>
            );
          })}
        </div>

        <div className="toolbar">
          <h2>{docTitle}</h2>
          <span className="hint">
            {current && current.blocks.length
              ? `${current.blocks.length} موضعًا به متشابهات`
              : ""}
          </span>
          <button
            className="btn"
            disabled={!current}
            onClick={() => current && downloadWord(current.num, current.data)}
          >
            ⬇️ تحميل ملف Word
          </button>
        </div>

        <div className="doc">
          {status === "loading" && (
            <div className="empty">
              …جارٍ تحميل بيانات سورة {SURAHS[active - 1]}
            </div>
          )}
          {status === "error" && (
            <div className="empty">تعذّر تحميل البيانات: {errorMsg}</div>
          )}
          {status === "idle" && !current && (
            <div className="empty">
              اضغط على رقم السورة لعرض المتشابهات وتوليد ملف Word.
            </div>
          )}
          {status === "idle" && current && <Preview blocks={current.blocks} />}
        </div>
      </div>
    </>
  );
}
