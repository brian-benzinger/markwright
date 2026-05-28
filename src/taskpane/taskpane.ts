import { markdownToOoxml } from "../convert";

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
  for (const el of document.querySelectorAll<HTMLElement>("body > [hidden], body > :not([hidden])")) {
    el.hidden = el.id !== id;
  }
}

function setStatus(message: string, kind: "info" | "error" = "info"): void {
  const status = document.getElementById("status");
  if (!status) return;
  status.textContent = message;
  status.classList.toggle("error", kind === "error");
}

async function onConvert(): Promise<void> {
  const input = document.getElementById("markdown-input") as HTMLTextAreaElement;
  const markdown = input.value;
  if (!markdown.trim()) {
    setStatus("Nothing to convert.", "error");
    return;
  }

  setStatus("Converting…");
  try {
    const ooxml = markdownToOoxml(markdown);
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
    setStatus("Done.");
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    setStatus(`Error: ${message}`, "error");
  }
}
