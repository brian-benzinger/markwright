import { parseMarkdown } from "../convert";
import {
  STYLE_TARGETS,
  type StyleMap,
  type StyleTarget,
  defaultStyleMap,
} from "../convert/styleMap";
import { applyBlocks } from "./apply";
import { loadDocumentParagraphStyles, loadStyleMap, saveStyleMap } from "./styles";

// Human-readable labels for each mappable construct, in the order they're
// shown in the panel (STYLE_TARGETS drives the order).
const TARGET_LABELS: Record<StyleTarget, string> = {
  heading1: "Heading 1",
  heading2: "Heading 2",
  heading3: "Heading 3",
  heading4: "Heading 4",
  heading5: "Heading 5",
  heading6: "Heading 6",
  paragraph: "Paragraph",
  blockquote: "Blockquote",
  codeBlock: "Code block",
};

// The active style map, loaded from the document's Settings on startup and
// passed to applyBlocks on every conversion.
let styleMap: StyleMap = defaultStyleMap();
// Whether the Styles panel has been populated yet (lazy on first open).
let stylesPopulated = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    show("unsupported");
    return;
  }
  show("app");

  styleMap = loadStyleMap();

  const convertBtn = document.getElementById("convert") as HTMLButtonElement;
  convertBtn.addEventListener("click", onConvert);

  const stylesToggle = document.getElementById("styles-toggle") as HTMLButtonElement;
  stylesToggle.addEventListener("click", onToggleStyles);

  const stylesSave = document.getElementById("styles-save") as HTMLButtonElement;
  stylesSave.addEventListener("click", onSaveStyles);
});

function show(id: string): void {
  for (const el of document.querySelectorAll<HTMLElement>("body > *")) {
    el.hidden = el.id !== id;
  }
}

let statusTimer: ReturnType<typeof setTimeout> | undefined;

function setStatus(message: string, kind: "info" | "error" = "info"): void {
  const status = document.getElementById("status");
  if (!status) return;
  if (statusTimer !== undefined) {
    clearTimeout(statusTimer);
    statusTimer = undefined;
  }
  status.textContent = message;
  status.classList.toggle("error", kind === "error");
}

// Success feedback that reports what landed in the document and then
// clears itself, so the pane doesn't sit on a stale "Done." forever.
function flashStatus(message: string): void {
  setStatus(message);
  statusTimer = setTimeout(() => {
    const status = document.getElementById("status");
    if (status) status.textContent = "";
    statusTimer = undefined;
  }, 2500);
}

async function onConvert(): Promise<void> {
  const input = document.getElementById("markdown-input") as HTMLTextAreaElement;
  const markdown = input.value;
  if (!markdown.trim()) {
    setStatus("Nothing to convert.", "error");
    return;
  }

  const blocks = parseMarkdown(markdown);
  if (blocks.length === 0) {
    setStatus("Nothing to convert.", "error");
    return;
  }

  setStatus("Converting…");
  try {
    await applyBlocks(blocks, styleMap);
    const n = blocks.length;
    flashStatus(`Inserted ${n} ${n === 1 ? "block" : "blocks"}.`);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    setStatus(`Error: ${message}`, "error");
  }
}

// Shows/hides the style-mapping panel, building its controls the first
// time it's opened (reading the host document's styles lazily).
async function onToggleStyles(): Promise<void> {
  const toggle = document.getElementById("styles-toggle") as HTMLButtonElement;
  const panel = document.getElementById("styles-panel") as HTMLElement;
  const willShow = panel.hidden;
  panel.hidden = !willShow;
  toggle.setAttribute("aria-expanded", String(willShow));
  if (willShow && !stylesPopulated) {
    stylesPopulated = true;
    await populateStylesPanel();
  }
}

async function populateStylesPanel(): Promise<void> {
  const grid = document.getElementById("styles-grid") as HTMLElement;
  setStylesStatus("Reading document styles…");
  let names: string[] = [];
  try {
    names = await loadDocumentParagraphStyles();
  } catch (err) {
    // A read failure shouldn't block mapping — the user can still pick
    // "(Default)" for everything. Surface it and carry on with no custom
    // options.
    const message = err instanceof Error ? err.message : String(err);
    setStylesStatus(`Couldn't read styles: ${message}`, "error");
  }
  grid.replaceChildren();
  for (const target of STYLE_TARGETS) {
    const select = buildSelect(target, names);
    const label = document.createElement("label");
    label.htmlFor = select.id;
    label.textContent = TARGET_LABELS[target];
    grid.append(label, select);
  }
  if (names.length === 0) {
    setStylesStatus("No named styles found in this document.");
  } else {
    setStylesStatus("");
  }
}

// Builds the <select> for one target: a leading "(Default)" option plus
// one option per host style name, preselecting the current mapping.
function buildSelect(target: StyleTarget, names: string[]): HTMLSelectElement {
  const select = document.createElement("select");
  select.id = `style-${target}`;
  select.dataset.target = target;

  const choice = styleMap[target];
  const current = "custom" in choice ? choice.custom : "";

  const def = document.createElement("option");
  def.value = "";
  def.textContent = "(Default)";
  select.append(def);

  // If the persisted custom name isn't in the document's style list (e.g.
  // a renamed or cross-locale style), keep it as a selectable option so
  // saving doesn't silently drop it.
  const options = current && !names.includes(current) ? [current, ...names] : names;
  for (const name of options) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    select.append(opt);
  }
  select.value = current;
  return select;
}

async function onSaveStyles(): Promise<void> {
  const grid = document.getElementById("styles-grid") as HTMLElement;
  const defaults = defaultStyleMap();
  const next: StyleMap = defaultStyleMap();
  for (const select of grid.querySelectorAll<HTMLSelectElement>("select[data-target]")) {
    const target = select.dataset.target as StyleTarget;
    next[target] = select.value ? { custom: select.value } : defaults[target];
  }
  setStylesStatus("Saving…");
  try {
    await saveStyleMap(next);
    styleMap = next;
    setStylesStatus("Saved.");
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    setStylesStatus(`Error: ${message}`, "error");
  }
}

function setStylesStatus(message: string, kind: "info" | "error" = "info"): void {
  const status = document.getElementById("styles-status");
  if (!status) return;
  status.textContent = message;
  status.classList.toggle("error", kind === "error");
}
