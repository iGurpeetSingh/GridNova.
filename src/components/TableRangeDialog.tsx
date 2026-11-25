import { useEffect, useState } from "react";
import { Table2, X } from "lucide-react";

interface Props {
  open: boolean;
  title: string;
  initialRange: string;
  initialHasHeaders: boolean;
  submitLabel: string;
  onClose: () => void;
  onSubmit: (range: string, hasHeaders: boolean) => void;
}
export default function TableRangeDialog({ open, title, initialRange, initialHasHeaders, submitLabel, onClose, onSubmit }: Props) {
  const [range, setRange] = useState(initialRange);
  const [hasHeaders, setHasHeaders] = useState(initialHasHeaders);
  useEffect(() => { if (open) { setRange(initialRange); setHasHeaders(initialHasHeaders); } }, [open, initialRange, initialHasHeaders]);
  if (!open) return null;
  return <div className="grid-dialog-backdrop" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) onClose(); }}>
    <section className="grid-dialog grid-table-range-dialog" role="dialog" aria-modal="true" aria-labelledby="grid-table-dialog-title">
      <header><div><span>GRIDNOVA TABLE</span><h2 id="grid-table-dialog-title">{title}</h2></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="grid-table-range-fields">
        <label>Where is the data for your table?
          <input autoFocus value={range} onChange={event => setRange(event.target.value)} placeholder="A1:D10" aria-label="Table data range"/>
        </label>
        <label className="grid-table-header-check"><input type="checkbox" checked={hasHeaders} onChange={event => setHasHeaders(event.target.checked)}/> My table has headers</label>
        <p>Choose a continuous cell range. Table headers must be unique; blank headers will be named automatically.</p>
      </div>
      <footer><button type="button" className="grid-dialog-secondary" onClick={onClose}>Cancel</button><button type="button" onClick={() => onSubmit(range, hasHeaders)}><Table2 size={14}/>{submitLabel}</button></footer>
    </section>
  </div>;
}
