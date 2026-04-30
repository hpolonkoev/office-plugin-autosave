export type LocaleStrings = {
  addin_name: string;
  activated_msg: string;
  toggle_on: string;
  toggle_off: string;
  last_saved: string;
  never_saved: string;
  interval_label: string;
  interval_unit: string;
  save_settings: string;
  settings_saved: string;
  unsaved_document: string;
  save_error: string;
  managed_by_it: string;
  protected_view: string;
};

let strings: LocaleStrings | null = null;

async function fetchLocale(locale: string): Promise<LocaleStrings | null> {
  try {
    const response = await fetch(`locales/${locale}.json`);
    if (!response.ok) return null;
    return (await response.json()) as LocaleStrings;
  } catch {
    return null;
  }
}

export async function initI18n(locale: string): Promise<void> {
  // Try exact locale (e.g. "fr-BE")
  strings = await fetchLocale(locale);
  if (strings) return;

  // Fall back to base language (e.g. "fr")
  const baseLang = locale.split("-")[0];
  if (baseLang !== locale) {
    strings = await fetchLocale(baseLang);
    if (strings) return;
  }

  // Final fallback to English
  strings = await fetchLocale("en");
}

export function t(key: keyof LocaleStrings): string {
  if (!strings) return key;
  return strings[key] ?? key;
}
