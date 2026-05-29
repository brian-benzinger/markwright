import { parseMarkdown, type Block, type Run } from "../convert";

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
  for (const el of document.querySelectorAll<HTMLElement>(
    "body > [hidden], body > :not([hidden])",
  )) {
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
  const input = document.getElementById(
    "markdown-input",
  ) as HTMLTextAreaElement;
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
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      // Clear any selected text, then insert the first block as a new
      // paragraph immediately before the cursor. Subsequent blocks chain
      // after that paragraph, so the original cursor position is preserved
      // at the end of the inserted content.
      selection.insertText("", Word.InsertLocation.replace);
      let para = selection.insertParagraph("", Word.InsertLocation.before);
      applyBlock(para, blocks[0]);
      for (let i = 1; i < blocks.length; i++) {
        para = para.insertParagraph("", Word.InsertLocation.after);
        applyBlock(para, blocks[i]);
      }
      await context.sync();
    });
    setStatus("Done.");
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    setStatus(`Error: ${message}`, "error");
  }
}

function applyBlock(paragraph: Word.Paragraph, block: Block): void {
  paragraph.styleBuiltIn =
    block.kind === "heading"
      ? headingStyle(block.level)
      : Word.BuiltInStyleName.normal;
  for (const run of block.runs) {
    applyRun(paragraph, run);
  }
}

function applyRun(paragraph: Word.Paragraph, run: Run): void {
  const range = paragraph.insertText(run.text, Word.InsertLocation.end);
  if (run.bold) range.font.bold = true;
  if (run.italic) range.font.italic = true;
  if (run.strike) range.font.strikeThrough = true;
  // Word lacks a built-in "code" character style; apply a monospace font directly.
  if (run.code) range.font.name = "Consolas";
  if (run.link) range.hyperlink = run.link;
}

function headingStyle(level: 1 | 2 | 3 | 4 | 5 | 6): Word.BuiltInStyleName {
  switch (level) {
    case 1:
      return Word.BuiltInStyleName.heading1;
    case 2:
      return Word.BuiltInStyleName.heading2;
    case 3:
      return Word.BuiltInStyleName.heading3;
    case 4:
      return Word.BuiltInStyleName.heading4;
    case 5:
      return Word.BuiltInStyleName.heading5;
    case 6:
      return Word.BuiltInStyleName.heading6;
  }
}
