import { useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { ReactNode, RefObject } from "react";

interface Props {
  open: boolean;
  anchorRef: RefObject<HTMLElement | null>;
  onClose: () => void;
  children: ReactNode;
  className?: string;
  align?: "left" | "right";
  role?: string;
  ariaLabel?: string;
}

export default function FloatingPortal({ open, anchorRef, onClose, children, className = "", align = "left", role, ariaLabel }: Props) {
  const panelRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const place = () => {
      const anchor = anchorRef.current, panel = panelRef.current;
      if (!anchor || !panel) return;
      const anchorRect = anchor.getBoundingClientRect(), panelRect = panel.getBoundingClientRect();
      const margin = 8, gap = 4;
      const preferredLeft = align === "right" ? anchorRect.right - panelRect.width : anchorRect.left;
      const left = Math.max(margin, Math.min(preferredLeft, window.innerWidth - panelRect.width - margin));
      const below = anchorRect.bottom + gap, above = anchorRect.top - panelRect.height - gap;
      const top = below + panelRect.height <= window.innerHeight - margin || above < margin
        ? Math.max(margin, Math.min(below, window.innerHeight - panelRect.height - margin))
        : above;
      setPosition(current => current?.left === left && current.top === top ? current : { left, top });
    };
    place();
    window.addEventListener("resize", place);
    window.addEventListener("scroll", place, true);
    return () => { window.removeEventListener("resize", place); window.removeEventListener("scroll", place, true); };
  }, [open, anchorRef, align, children]);

  useLayoutEffect(() => {
    if (!open) { setPosition(null); return; }
    const closeOutside = (event: PointerEvent) => {
      const target = event.target as Node;
      if (!panelRef.current?.contains(target) && !anchorRef.current?.contains(target)) onClose();
    };
    const closeEscape = (event: KeyboardEvent) => { if (event.key === "Escape") { event.preventDefault(); onClose(); } };
    document.addEventListener("pointerdown", closeOutside);
    document.addEventListener("keydown", closeEscape);
    return () => { document.removeEventListener("pointerdown", closeOutside); document.removeEventListener("keydown", closeEscape); };
  }, [open, anchorRef, onClose]);

  if (!open) return null;
  return createPortal(
    <div ref={panelRef} role={role} aria-label={ariaLabel} className={`floating-portal ${className}`} style={{ left: position?.left ?? -10000, top: position?.top ?? -10000 }}>
      {children}
    </div>,
    document.body,
  );
}
