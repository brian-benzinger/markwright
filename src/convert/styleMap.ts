// User-configurable mapping from Markdown constructs to Word paragraph
// styles (Milestone 5). This module is deliberately pure — it never
// touches Office.js — so it can be unit-tested under vitest and stays in
// coverage scope. The applier (src/taskpane/apply.ts) translates a
// StyleChoice into the actual Word API call; the task pane persists the
// map via the Office Settings API.

// The Markdown constructs whose paragraph style can be remapped. List
// items are intentionally absent: the applier must not set a built-in
// style on a list paragraph (it detaches list membership), so they're
// out of scope for the mapping UI.
export type StyleTarget =
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "heading5"
  | "heading6"
  | "paragraph"
  | "blockquote"
  | "codeBlock";

// Locale-invariant identifiers for the Word built-in styles we map to by
// default. The applier turns these into Word.BuiltInStyleName values, so
// the zero-config path stays byte-for-byte identical to the pre-M5
// behaviour and keeps working in non-English Word installs.
export type StyleToken =
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "heading5"
  | "heading6"
  | "normal"
  | "quote";

// A resolved style binding. `builtIn` carries a locale-invariant token
// (used by defaults, applied via styleBuiltIn). `custom` carries a host
// document style's localised name read live from getStyles(), applied via
// paragraph.style — correct for the install it was chosen on.
export type StyleChoice = { builtIn: StyleToken } | { custom: string };

export type StyleMap = Record<StyleTarget, StyleChoice>;

// All targets in display order. Single source of truth for the UI and for
// validation/merge so the two never drift.
export const STYLE_TARGETS: readonly StyleTarget[] = [
  "heading1",
  "heading2",
  "heading3",
  "heading4",
  "heading5",
  "heading6",
  "paragraph",
  "blockquote",
  "codeBlock",
];

const VALID_TOKENS: ReadonlySet<string> = new Set<StyleToken>([
  "heading1",
  "heading2",
  "heading3",
  "heading4",
  "heading5",
  "heading6",
  "normal",
  "quote",
]);

// The default binding for each target, reproducing the original hard-coded
// applier behaviour: headings -> Heading 1..6, blockquote -> Quote, and
// paragraph / code block -> Normal.
export function defaultStyleMap(): StyleMap {
  return {
    heading1: { builtIn: "heading1" },
    heading2: { builtIn: "heading2" },
    heading3: { builtIn: "heading3" },
    heading4: { builtIn: "heading4" },
    heading5: { builtIn: "heading5" },
    heading6: { builtIn: "heading6" },
    paragraph: { builtIn: "normal" },
    blockquote: { builtIn: "quote" },
    codeBlock: { builtIn: "normal" },
  };
}

// Narrows an unknown (e.g. JSON parsed out of the Settings store) into a
// StyleChoice, or null if it isn't one. Built-in tokens must be from the
// known set; custom names must be non-empty strings.
function asStyleChoice(value: unknown): StyleChoice | null {
  if (typeof value !== "object" || value === null) return null;
  if ("builtIn" in value) {
    const token = (value as { builtIn: unknown }).builtIn;
    return typeof token === "string" && VALID_TOKENS.has(token)
      ? { builtIn: token as StyleToken }
      : null;
  }
  if ("custom" in value) {
    const name = (value as { custom: unknown }).custom;
    return typeof name === "string" && name.length > 0 ? { custom: name } : null;
  }
  return null;
}

// Merges a persisted (and therefore untrusted) value over the defaults.
// Any target whose stored entry is missing or malformed falls back to its
// default, so a partial or corrupted Settings payload can never produce an
// invalid map.
export function mergeStyleMap(persisted: unknown): StyleMap {
  const map = defaultStyleMap();
  if (typeof persisted !== "object" || persisted === null) return map;
  const record = persisted as Record<string, unknown>;
  for (const target of STYLE_TARGETS) {
    const choice = asStyleChoice(record[target]);
    if (choice) map[target] = choice;
  }
  return map;
}
