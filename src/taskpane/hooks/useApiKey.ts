import { useState, useEffect, useCallback } from "react";

const STORAGE_KEY_API = "vsme_api_key";
const STORAGE_KEY_URL = "vsme_api_url";
const DEFAULT_API_URL = "https://app.globalmoo.com/api/";

// Persistence helpers — try OfficeRuntime.storage first, fall back to localStorage
async function loadItem(key: string): Promise<string | null> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const val = await OfficeRuntime.storage.getItem(key);
      if (val) return val;
    }
  } catch {
    // OfficeRuntime.storage unavailable or blocked
  }
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

async function saveItem(key: string, value: string): Promise<void> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(key, value);
    }
  } catch {
    // OfficeRuntime.storage unavailable
  }
  try {
    localStorage.setItem(key, value);
  } catch {
    // localStorage unavailable
  }
}

async function removeItem(key: string): Promise<void> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.removeItem(key);
    }
  } catch {
    // ignore
  }
  try {
    localStorage.removeItem(key);
  } catch {
    // ignore
  }
}

export function useApiKey() {
  const [apiKey, setApiKeyState] = useState<string>("");
  const [apiUrl, setApiUrlState] = useState<string>(DEFAULT_API_URL);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    (async () => {
      const savedKey = await loadItem(STORAGE_KEY_API);
      const savedUrl = await loadItem(STORAGE_KEY_URL);
      if (savedKey) setApiKeyState(savedKey);
      if (savedUrl) setApiUrlState(savedUrl);
      setIsLoading(false);
    })();
  }, []);

  const setApiKey = useCallback(async (key: string) => {
    setApiKeyState(key);
    if (key) {
      await saveItem(STORAGE_KEY_API, key);
    } else {
      await removeItem(STORAGE_KEY_API);
    }
  }, []);

  const setApiUrl = useCallback(async (url: string) => {
    const normalized = url.trim() || DEFAULT_API_URL;
    setApiUrlState(normalized);
    await saveItem(STORAGE_KEY_URL, normalized);
  }, []);

  const clearApiKey = useCallback(async () => {
    setApiKeyState("");
    await removeItem(STORAGE_KEY_API);
  }, []);

  return { apiKey, apiUrl, isLoading, setApiKey, setApiUrl, clearApiKey, DEFAULT_API_URL };
}
