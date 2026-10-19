import { useEffect, useReducer, useRef, useState } from "react";
import { initialWorkspace, reducer } from "../state/workspace";
import { createWorkspaceSaver, readLocalWorkspace, writeLocalWorkspace, type SaveState } from "../services/localWorkspace";

export function useLocalWorkspace() {
  const [workspace, dispatch] = useReducer(reducer, undefined, initialWorkspace);
  const [ready, setReady] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [saveState, setSaveState] = useState<SaveState>({ status: "saved" });
  const saver = useRef<ReturnType<typeof createWorkspaceSaver> | null>(null);
  useEffect(() => {
    let active = true;
    void readLocalWorkspace().then(saved => {
      if (!active) return;
      const restored = saved?.workspace ?? initialWorkspace();
      saver.current = createWorkspaceSaver(restored, saved?.revision ?? null, writeLocalWorkspace,
        state => { if (active) setSaveState(state); });
      dispatch({ type: "restore", workspace: restored });
      setSaveState({ status: "saved", savedAt: saved?.savedAt });
      setReady(true);
    }).catch(error => { if (active) setLoadError((error as Error).message); });
    return () => { active = false; };
  }, []);
  useEffect(() => {
    if (ready) saver.current?.update(workspace);
  }, [workspace, ready]);
  useEffect(() => {
    const onLeave = (event: BeforeUnloadEvent) => {
      if (saver.current?.hasPending()) { event.preventDefault(); event.returnValue = ""; }
    };
    window.addEventListener("beforeunload", onLeave);
    return () => window.removeEventListener("beforeunload", onLeave);
  }, []);
  return { workspace, dispatch, ready, loadError, saveState, retrySave: () => void saver.current?.retry() };
}
