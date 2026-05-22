// Persists WorkbookState to Excel custom XML parts with Workbook.settings fallback
import { WorkbookState, DEFAULT_WORKBOOK_STATE, STATE_SCHEMA_VERSION } from "../types/workbookState";

/**
 * Apply forward migrations to a persisted state. v1 → v2 shifts wizardStep up
 * by 1 to account for the Resume step inserted at index 1.
 */
function migrateState(parsed: Partial<WorkbookState>): Partial<WorkbookState> {
  const version = parsed.schemaVersion ?? 1;
  if (version >= STATE_SCHEMA_VERSION) return parsed;
  const next: Partial<WorkbookState> = { ...parsed };
  if (version < 2 && typeof next.wizardStep === "number" && next.wizardStep > 0) {
    next.wizardStep = next.wizardStep + 1;
  }
  next.schemaVersion = STATE_SCHEMA_VERSION;
  return next;
}

const XML_NAMESPACE = "gmoo-globalmoo-state";
const SETTINGS_KEY_PREFIX = "gmoo_";

// Keys we also store in Workbook.settings as fallback (critical IDs).
// schemaVersion is included so migrateState() can correctly tell whether a
// fallback-restored wizardStep has already been migrated. Without it the
// migration would shift v2 wizardSteps a second time on every fallback load.
const FALLBACK_KEYS: (keyof WorkbookState)[] = [
  "schemaVersion",
  "modelId",
  "projectId",
  "trialId",
  "objectiveId",
  "wizardStep",
  "apiKeyHint",
  "activeConnectionId",
];

export async function saveWorkbookState(state: WorkbookState): Promise<void> {
  await Excel.run(async (context) => {
    // Save full state to custom XML part
    const xmlParts = context.workbook.customXmlParts;
    const existingParts = xmlParts.getByNamespace(XML_NAMESPACE);
    existingParts.load("items");
    await context.sync();

    // Delete existing parts for our namespace
    for (const part of existingParts.items) {
      part.delete();
    }

    const xmlContent = `<gmooState xmlns="${XML_NAMESPACE}">${escapeXml(JSON.stringify(state))}</gmooState>`;
    xmlParts.add(xmlContent);

    // Also save critical IDs to Workbook.settings as fallback
    const settings = context.workbook.settings;
    for (const key of FALLBACK_KEYS) {
      const value = state[key];
      settings.add(`${SETTINGS_KEY_PREFIX}${key}`, value != null ? String(value) : "");
    }

    await context.sync();
  });
}

export async function loadWorkbookState(): Promise<WorkbookState> {
  return Excel.run(async (context) => {
    // Try custom XML parts first
    const xmlParts = context.workbook.customXmlParts.getByNamespace(XML_NAMESPACE);
    xmlParts.load("items");
    await context.sync();

    if (xmlParts.items.length > 0) {
      const xmlPart = xmlParts.items[0];
      const xml = xmlPart.getXml();
      await context.sync();

      try {
        const jsonStr = unescapeXml(
          xml.value.replace(/<\/?gmooState[^>]*>/g, "")
        );
        const parsed = JSON.parse(jsonStr) as Partial<WorkbookState>;
        const migrated = migrateState(parsed);
        return { ...DEFAULT_WORKBOOK_STATE, ...migrated };
      } catch {
        // XML part corrupted, fall through to settings fallback
      }
    }

    // Fallback: try to restore critical IDs from Workbook.settings
    const settings = context.workbook.settings;
    const restored: Partial<WorkbookState> = {};
    let hasAny = false;

    for (const key of FALLBACK_KEYS) {
      const setting = settings.getItemOrNullObject(`${SETTINGS_KEY_PREFIX}${key}`);
      setting.load("value");
    }
    await context.sync();

    for (const key of FALLBACK_KEYS) {
      const setting = settings.getItemOrNullObject(`${SETTINGS_KEY_PREFIX}${key}`);
      if (!setting.isNullObject && setting.value) {
        hasAny = true;
        const val = setting.value;
        if (
          key === "wizardStep" ||
          key === "modelId" ||
          key === "projectId" ||
          key === "trialId" ||
          key === "objectiveId" ||
          key === "schemaVersion"
        ) {
          const num = parseInt(val, 10);
          if (!isNaN(num)) {
            (restored as Record<string, unknown>)[key] = num;
          }
        } else {
          (restored as Record<string, unknown>)[key] = val;
        }
      }
    }

    if (hasAny) {
      // restored carries its own schemaVersion (if previously saved) so
      // migrateState can skip the v1→v2 wizardStep shift when not needed.
      return { ...DEFAULT_WORKBOOK_STATE, ...migrateState(restored) };
    }

    return { ...DEFAULT_WORKBOOK_STATE };
  });
}

export async function clearWorkbookState(): Promise<void> {
  await Excel.run(async (context) => {
    const xmlParts = context.workbook.customXmlParts.getByNamespace(XML_NAMESPACE);
    xmlParts.load("items");
    await context.sync();

    for (const part of xmlParts.items) {
      part.delete();
    }

    const settings = context.workbook.settings;
    for (const key of FALLBACK_KEYS) {
      const setting = settings.getItemOrNullObject(`${SETTINGS_KEY_PREFIX}${key}`);
      setting.load("isNullObject");
    }
    await context.sync();

    for (const key of FALLBACK_KEYS) {
      const setting = settings.getItemOrNullObject(`${SETTINGS_KEY_PREFIX}${key}`);
      if (!setting.isNullObject) {
        setting.delete();
      }
    }
    await context.sync();
  });
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeXml(str: string): string {
  return str
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}
