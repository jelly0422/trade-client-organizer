import { useEffect } from "react";

const selector =
  ".ant-btn:not(.ant-btn-link), .nav-item, .category-card, .file-chip, .ant-drawer-close, .ant-modal-close";

/** Track one hovered control without rerendering the workspace or grid. */
export function useGlassPointer() {
  useEffect(() => {
    const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");
    let active: HTMLElement | null = null;
    let frame = 0;
    let x = 0;
    let y = 0;
    const clear = () => {
      cancelAnimationFrame(frame);
      frame = 0;
      active?.style.removeProperty("--glass-x");
      active?.style.removeProperty("--glass-y");
      active = null;
    };
    const move = (event: PointerEvent) => {
      if (event.pointerType === "touch" || reducedMotion.matches) return;
      const control = event.target instanceof Element
        ? event.target.closest<HTMLElement>(selector)
        : null;
      if (!control || control.matches(":disabled, [aria-disabled='true']")) {
        clear();
        return;
      }
      if (control !== active) {
        clear();
        active = control;
      }
      x = event.clientX;
      y = event.clientY;
      if (!frame) frame = requestAnimationFrame(() => {
        frame = 0;
        if (!active?.isConnected) return clear();
        const bounds = active.getBoundingClientRect();
        active.style.setProperty("--glass-x", `${x - bounds.left}px`);
        active.style.setProperty("--glass-y", `${y - bounds.top}px`);
      });
    };
    const leave = (event: PointerEvent) => {
      if (!(event.relatedTarget instanceof Node) || !active?.contains(event.relatedTarget)) clear();
    };
    const focus = () => clear();
    document.addEventListener("pointermove", move, { passive: true });
    document.addEventListener("pointerout", leave, { passive: true });
    document.addEventListener("focusin", focus);
    window.addEventListener("blur", clear);
    reducedMotion.addEventListener("change", clear);
    return () => {
      clear();
      document.removeEventListener("pointermove", move);
      document.removeEventListener("pointerout", leave);
      document.removeEventListener("focusin", focus);
      window.removeEventListener("blur", clear);
      reducedMotion.removeEventListener("change", clear);
    };
  }, []);
}
