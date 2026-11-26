import { useEffect, useMemo, useState } from "react";
import { FunctionSquare, Search, X } from "lucide-react";
import { FORMULA_CATALOG, FORMULA_CATEGORIES, IMPLEMENTED_FORMULA_COUNT } from "../lib/formula-catalog";

interface Props { open: boolean; onClose: () => void; onInsert: (formula: string) => void; initialCategory?:string; }
export default function FormulaLibraryDialog({ open, onClose, onInsert, initialCategory="All" }: Props) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState<string>("All");
  const [supportedOnly, setSupportedOnly] = useState(false);
  useEffect(()=>{if(open)setCategory(FORMULA_CATEGORIES.includes(initialCategory as typeof FORMULA_CATEGORIES[number])?initialCategory:"All");},[open,initialCategory]);
  const results = useMemo(() => FORMULA_CATALOG.filter(entry =>
    (category === "All" || entry.category === category) &&
    (!supportedOnly || entry.implemented) &&
    (!query.trim() || `${entry.name} ${entry.description} ${entry.syntax}`.toLowerCase().includes(query.toLowerCase()))
  ), [category, query, supportedOnly]);
  if (!open) return null;
  return <div className="grid-dialog-backdrop formula-library-backdrop" role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) onClose(); }}>
    <section className="grid-dialog formula-library-dialog" role="dialog" aria-modal="true" aria-labelledby="formula-library-title">
      <header><div><span>FUNCTION LIBRARY</span><h2 id="formula-library-title">Insert Function</h2><p>{IMPLEMENTED_FORMULA_COUNT} available in the current calculation engine · {FORMULA_CATALOG.length} catalogue entries</p></div><button type="button" onClick={onClose} aria-label="Close"><X size={18}/></button></header>
      <div className="formula-library-toolbar">
        <label className="formula-library-search"><Search size={15}/><input autoFocus value={query} onChange={event => setQuery(event.target.value)} placeholder="Search a function or describe what you want to do"/></label>
        <label className="formula-library-supported"><input type="checkbox" checked={supportedOnly} onChange={event => setSupportedOnly(event.target.checked)}/> Engine-supported only</label>
      </div>
      <nav className="formula-library-categories" aria-label="Function categories">
        <button type="button" className={category === "All" ? "active" : ""} onClick={() => setCategory("All")}>All</button>
        {FORMULA_CATEGORIES.map(item => <button type="button" key={item} className={category === item ? "active" : ""} onClick={() => setCategory(item)}>{item}</button>)}
      </nav>
      <div className="formula-library-results">
        {results.map(entry => <article key={entry.name} className="formula-library-entry">
          <div className="formula-library-entry-icon"><FunctionSquare size={17}/></div>
          <div className="formula-library-entry-main"><div className="formula-library-entry-heading"><strong>{entry.name}</strong><span>{entry.category}</span></div><code>{entry.syntax}</code><p>{entry.description}</p><small>Example: {entry.example}</small></div>
          <div className="formula-library-entry-action">{entry.implemented ? <span className="formula-supported">Supported</span> : <span className="formula-catalogue-only">Catalogue</span>}<button type="button" onClick={() => { onInsert(`=${entry.name}()`); onClose(); }}>Insert</button></div>
        </article>)}
        {!results.length && <div className="formula-library-empty">No matching functions. Try another name or category.</div>}
      </div>
      <footer><span>Catalogue entries are references; catalogue-only functions are not yet evaluated by GridNova.</span><button type="button" onClick={onClose}>Close</button></footer>
    </section>
  </div>;
}
