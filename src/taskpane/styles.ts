import { mergeStyleMap, type StyleMap } from "../convert/styleMap";

// Office.js glue for Milestone 5: reading the host document's paragraph
// styles and persisting the user's style map. Kept out of src/convert/
// (and so out of coverage scope) because it can only run inside Word —
// the pure mapping logic lives in src/convert/styleMap.ts.

const SETTINGS_KEY = "markwright:styleMap";

// Reads the active document's paragraph styles and returns their
// localised names, sorted, deduplicated. Returns an empty list if the
// host is too old to support getStyles (WordApi 1.5) so the caller can
// degrade gracefully rather than throw.
export async function loadDocumentParagraphStyles(): Promise<string[]> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) return [];
  return Word.run(async (context) => {
    const styles = context.document.getStyles();
    styles.load("items/nameLocal,items/type");
    await context.sync();
    const names = styles.items
      .filter((s) => s.type === Word.StyleType.paragraph && s.nameLocal)
      .map((s) => s.nameLocal);
    return [...new Set(names)].sort((a, b) => a.localeCompare(b));
  });
}

// Reads the persisted style map from the document's Settings, merged over
// the defaults so a missing or corrupt payload yields a valid map.
export function loadStyleMap(): StyleMap {
  const raw = Office.context.document.settings.get(SETTINGS_KEY) as unknown;
  return mergeStyleMap(raw);
}

// Persists the style map to the document's Settings. The set is in-memory
// until saveAsync flushes it to the document; we promisify that callback.
export function saveStyleMap(map: StyleMap): Promise<void> {
  Office.context.document.settings.set(SETTINGS_KEY, map);
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}
