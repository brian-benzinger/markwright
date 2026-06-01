import { parseMarkdown } from "../convert";
import { applyBlocks } from "./apply";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    show("unsupported");
    return;
  }
  show("app");

  const convertBtn = document.getElementById("convert") as HTMLButtonElement;
  convertBtn.addEventListener("click", onConvert);
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
    await applyBlocks(blocks);
    const n = blocks.length;
    flashStatus(`Inserted ${n} ${n === 1 ? "block" : "blocks"}.`);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    setStatus(`Error: ${message}`, "error");
  }
}
