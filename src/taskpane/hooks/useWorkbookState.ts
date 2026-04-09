import { useState, useCallback, useEffect } from "react";
import { WorkbookState, DEFAULT_WORKBOOK_STATE } from "../types/workbookState";
import { saveWorkbookState, loadWorkbookState, clearWorkbookState } from "../services/stateStore";

export function useWorkbookState() {
  const [state, setState] = useState<WorkbookState>(DEFAULT_WORKBOOK_STATE);
  const [isLoaded, setIsLoaded] = useState(false);

  // Load state on mount
  useEffect(() => {
    loadState();
  }, []);

  const loadState = async () => {
    try {
      const loaded = await loadWorkbookState();
      setState(loaded);
    } catch {
      // Use defaults if loading fails
    } finally {
      setIsLoaded(true);
    }
  };

  const updateState = useCallback(async (updates: Partial<WorkbookState>) => {
    setState((prev) => {
      const next = { ...prev, ...updates };
      // Fire-and-forget save
      saveWorkbookState(next).catch(() => {});
      return next;
    });
  }, []);

  const resetState = useCallback(async () => {
    setState(DEFAULT_WORKBOOK_STATE);
    try {
      await clearWorkbookState();
    } catch {
      // Ignore errors
    }
  }, []);

  return { state, isLoaded, updateState, resetState };
}
