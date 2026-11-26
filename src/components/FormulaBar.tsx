import { ChevronDown, FunctionSquare } from "lucide-react";

interface Props {
  address: string;
  value: string;
  onChange: (value: string) => void;
  onCommit: () => void;
  readOnly?: boolean;
}

export default function FormulaBar({ address, value, onChange, onCommit, readOnly = false }: Props) {
  return (
    <div className={`formula-bar no-print ${readOnly ? "formula-bar-readonly" : ""}`}>
      <button type="button" onMouseDown={event => event.preventDefault()} className="formula-name-box" aria-label="Name box" title="Name box">
        {address}<ChevronDown size={11}/>
      </button>
      <span className="formula-divider" aria-hidden="true"/>
      <button type="button" disabled={readOnly} onMouseDown={event => event.preventDefault()} className="formula-control formula-cancel" title="Cancel" aria-label="Cancel" onClick={() => onChange("")}>×</button>
      <button type="button" disabled={readOnly} onMouseDown={event => event.preventDefault()} className="formula-control formula-enter" title="Enter" aria-label="Enter" onClick={onCommit}>✓</button>
      <button type="button" onMouseDown={event => event.preventDefault()} className="formula-fx" title="Insert Function" aria-label="Insert Function">
        <FunctionSquare size={15}/><i>fx</i>
      </button>
      <input
        aria-label="Formula bar input"
        value={value}
        readOnly={readOnly}
        onChange={event => onChange(event.target.value)}
        onKeyDown={event => {
          if (!readOnly && event.key === "Enter" && !event.nativeEvent.isComposing && event.keyCode !== 229) onCommit();
        }}
        className="formula-input"
      />
      <button type="button" onMouseDown={event => event.preventDefault()} className="formula-expand" title="Expand formula bar" aria-label="Expand formula bar">
        <ChevronDown size={12}/>
      </button>
    </div>
  );
}
