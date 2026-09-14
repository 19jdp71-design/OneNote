# Markdown for OneNote

A OneNote task-pane add-in: type Markdown in a side panel, see it live-rendered,
and insert it into the current page as real OneNote content (headings, bold/italic,
inline code, links, blockquotes, numbered/bulleted lists, and native checkboxes).

This is deliberately built with **zero build tooling** — no npm, no webpack, no
`office-addin-dev-certs`. It's a static `taskpane.html` + `taskpane.js` + `taskpane.css`
loaded straight into OneNote over HTTPS. That was a hard requirement: `office-addin-dev-certs`
is blocked by group policy on the ATEC machine, so this add-in never needs it, on
any machine.

## How it works

- **Left pane** — a plain `<textarea>` where you type Markdown.
- **Right pane** — a live preview, rendered in the browser with [marked.js](https://marked.js.org/).
- **Insert into Page** — converts the Markdown into the specific HTML subset OneNote's
  page model understands (via a second, custom marked.js renderer) and inserts it as a
  new outline on the current page using the OneNote JavaScript API (`page.addOutline`).
- Checkboxes (`- [ ]` / `- [x]`) become native OneNote to-do paragraphs
  (`<p data-tag="to-do">` / `data-tag="to-do:completed"`), not plain `<input>` checkboxes.

**Known limitation (v1):** insertion is one-way (Markdown → OneNote). Reading an
existing OneNote page back into Markdown is not implemented — OneNote's page HTML is
complex enough that a faithful reverse conversion is its own project. `addOutline`
always drops the new content at a fixed spot (100, 100) on the page; drag it wherever
you like afterward.

## Project layout

```
manifest.xml       classic Office Add-in XML manifest (Hosts: Note)
taskpane.html       the task pane UI
taskpane.css
taskpane.js         markdown rendering + OneNote insertion logic
assets/             icon-16.png / icon-32.png / icon-80.png (placeholders — swap freely)
```

## 1. Host it somewhere HTTPS (required, one-time)

OneNote add-ins must load their task pane over HTTPS. We use **GitHub Pages** because
it's free, requires no server maintenance, and needs nothing running locally on either
machine — so it works the same way on your personal PC and on the locked-down ATEC box.

1. This project is wired to **github.com/19jdp71-design/OneNote**. Push it there:

   ```bash
   git init
   git add .
   git commit -m "Initial Markdown for OneNote add-in"
   git branch -M main
   git remote add origin https://github.com/19jdp71-design/OneNote.git
   git push -u origin main
   ```

2. On GitHub: **Settings → Pages → Source → Deploy from branch → `main` / `/ (root)`** → Save.
   This publishes to `https://19jdp71-design.github.io/OneNote/`, which `manifest.xml`
   already points at.

3. Wait a minute for Pages to publish, then confirm `taskpane.html` loads by visiting
   `https://19jdp71-design.github.io/OneNote/taskpane.html` in a browser.

   If you ever rename the repo or use a different account, update the 8 URLs in
   `manifest.xml` (`IconUrl`, `HighResolutionIconUrl`, `SupportUrl`, `AppDomain`,
   `SourceLocation`, the 3 image resources, and the taskpane URL resource) to match.

## 2. Sideload into OneNote (Windows desktop)

No npm, no `npm start`, no dev certs — just point OneNote at the manifest:

1. Open OneNote (desktop app) → **Insert** tab → **Add-ins** → **My Add-ins**.
2. Click the small dropdown/gear icon → **Upload My Add-in**.
3. Browse to `manifest.xml` in this folder and select it.
4. A **Markdown** group with a **Markdown Editor** button appears on the **Home** tab.
   Click it to open the task pane.

Repeat this on each machine (personal PC and, if allowed, the ATEC machine) — sideloading
only requires OneNote to reach your GitHub Pages URL, no local install step.

**ATEC machine caveat:** this assumes `github.io` isn't blocked by the network content
filter and that "Upload My Add-in" sideloading itself isn't disabled by policy. If either
is blocked, the fallback is asking your OneNote/M365 admin to add it via the admin center's
"Centralized Deployment," which doesn't need sideloading permissions at all — but that's
an IT request, not something to solve from the client.

## Iterating on the code

Since there's no build step, you can test the Markdown → HTML logic directly in a normal
browser tab without touching OneNote at all — just open `taskpane.html` from a local static
server (e.g. VS Code's "Live Server" extension, or `python -m http.server` from this folder)
and it'll show a "not running inside OneNote" banner with the preview pane still fully
working. Once the rendering looks right, push to GitHub and click **Insert into Page**
sideloaded inside OneNote to test the real insertion.

To change anything about how Markdown maps to OneNote's HTML (fonts, checkbox styling,
heading sizes), edit the `createOneNoteRenderer()` function in `taskpane.js`.
