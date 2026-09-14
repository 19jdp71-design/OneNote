/* global Office, OneNote, marked, DOMPurify */

const mdInput = document.getElementById("md-input");
const preview = document.getElementById("preview");
const btnInsert = document.getElementById("btn-insert");
const btnCopy = document.getElementById("btn-copy");
const status = document.getElementById("status");
const offlineBanner = document.getElementById("offline-banner");

let runningInOffice = false;

// Renders Markdown for on-screen preview (normal HTML, browser-rendered).
function renderPreview(mdText) {
  const rawHtml = marked.parse(mdText, { gfm: true, breaks: false });
  preview.innerHTML = DOMPurify.sanitize(rawHtml);
}

function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Renders Markdown into the HTML subset OneNote's page model understands:
// plain <p>/<b>/<i>/<a>/<ul>/<ol>/<li>/<table>, and data-tag paragraphs for checkboxes.
function createOneNoteRenderer() {
  const renderer = new marked.Renderer();

  const headingSizes = { 1: "22pt", 2: "18pt", 3: "15pt", 4: "13pt", 5: "11pt", 6: "10pt" };
  renderer.heading = (text, level) => {
    const size = headingSizes[level] || "11pt";
    return `<p style="font-weight:bold;font-size:${size}">${text}</p>`;
  };

  renderer.paragraph = (text) => `<p>${text}</p>`;
  renderer.strong = (text) => `<b>${text}</b>`;
  renderer.em = (text) => `<i>${text}</i>`;
  renderer.del = (text) => `<span style="text-decoration:line-through">${text}</span>`;
  renderer.codespan = (code) =>
    `<span style="font-family:Consolas,monospace;background:#f0f0f0">${code}</span>`;

  renderer.code = (code) => {
    return code
      .split("\n")
      .map((line) => {
        const escaped = escapeHtml(line);
        return `<p style="font-family:Consolas,monospace;background:#f4f4f4;margin:0">${escaped || "&nbsp;"}</p>`;
      })
      .join("");
  };

  renderer.blockquote = (quote) =>
    `<div style="border-left:3px solid #cccccc;padding-left:8px;color:#666666">${quote}</div>`;

  renderer.hr = () => `<p style="border-bottom:1px solid #cccccc">&nbsp;</p>`;

  renderer.link = (href, _title, text) => `<a href="${href}">${text}</a>`;

  renderer.list = (body, ordered) => {
    // Task-list items render as OneNote checkbox paragraphs, which must NOT be
    // wrapped in <ul>/<ol> — OneNote's data-tag checkboxes are top-level paragraphs.
    if (body.includes('data-tag="to-do')) {
      return body;
    }
    return ordered ? `<ol>${body}</ol>` : `<ul>${body}</ul>`;
  };

  // marked injects a raw <input type="checkbox"> into the item text before this
  // renderer runs; suppress it (via renderer.checkbox below) so we control the
  // checkbox purely through OneNote's own data-tag attribute.
  renderer.checkbox = () => "";

  renderer.listitem = (text, task, checked) => {
    if (task) {
      const tag = checked ? "to-do:completed" : "to-do";
      return `<p data-tag="${tag}">${text.trim()}</p>`;
    }
    return `<li>${text}</li>`;
  };

  renderer.table = (header, body) => `<table><thead>${header}</thead><tbody>${body}</tbody></table>`;
  renderer.tablerow = (content) => `<tr>${content}</tr>`;
  renderer.tablecell = (content, flags) => (flags.header ? `<th>${content}</th>` : `<td>${content}</td>`);

  return renderer;
}

const oneNoteRenderer = createOneNoteRenderer();

function markdownToOneNoteHtml(mdText) {
  return marked.parse(mdText, { gfm: true, breaks: false, renderer: oneNoteRenderer });
}

function setStatus(text, isError) {
  status.textContent = text;
  status.style.color = isError ? "#a4262c" : "#666";
  if (text) {
    setTimeout(() => {
      if (status.textContent === text) status.textContent = "";
    }, 4000);
  }
}

async function insertIntoPage() {
  const mdText = mdInput.value;
  if (!mdText.trim()) {
    setStatus("Nothing to insert.", true);
    return;
  }

  const html = markdownToOneNoteHtml(mdText);

  try {
    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
      page.addOutline(100, 100, html);
      await context.sync();
    });
    setStatus("Inserted into page.");
  } catch (err) {
    console.error(err);
    setStatus("Insert failed: " + (err && err.message ? err.message : err), true);
  }
}

async function copyMarkdown() {
  try {
    await navigator.clipboard.writeText(mdInput.value);
    setStatus("Markdown copied.");
  } catch (err) {
    setStatus("Copy failed: " + err.message, true);
  }
}

let debounceTimer;
mdInput.addEventListener("input", () => {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => renderPreview(mdInput.value), 150);
});

btnCopy.addEventListener("click", copyMarkdown);
btnInsert.addEventListener("click", insertIntoPage);

renderPreview(mdInput.value);

if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
      runningInOffice = true;
      btnInsert.disabled = false;
    } else {
      btnInsert.disabled = true;
      offlineBanner.hidden = false;
    }
  });
} else {
  btnInsert.disabled = true;
  offlineBanner.hidden = false;
}
