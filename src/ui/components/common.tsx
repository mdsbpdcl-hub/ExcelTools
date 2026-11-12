import { useEffect, useId, useRef, useState, type ReactNode } from 'react';
import type { CellValue } from '../../core/types';
import { cellToString } from '../../core/util';

export function Stepper({ steps, current, maxReached, onGo }: { steps: string[]; current: number; maxReached: number; onGo: (i: number) => void }) {
  return (
    <ol className="stepper" aria-label="Steps">
      {steps.map((s, i) => (
        <li key={s} className={i < current ? 'done' : i === current ? 'current' : ''}>
          <button type="button" disabled={i > maxReached} onClick={() => onGo(i)} aria-current={i === current ? 'step' : undefined}>
            <span className="step-dot">{i < current ? '✓' : i + 1}</span>
            <span className="step-label">{s}</span>
          </button>
          {i < steps.length - 1 && <span className="step-line" />}
        </li>
      ))}
    </ol>
  );
}

export function PageHead({ title, sub }: { title: string; sub?: string }) {
  return (
    <div className="page-head">
      <h1>{title}</h1>
      {sub && <p>{sub}</p>}
    </div>
  );
}

export function Tiles({ items }: { items: { label: string; value: ReactNode; tone?: 'ok' | 'warn' | 'bad' }[] }) {
  return (
    <div className="tiles">
      {items.map((t) => (
        <div key={t.label} className={`tile ${t.tone ?? ''}`}>
          <div className="num">{t.value}</div>
          <div className="lbl">{t.label}</div>
        </div>
      ))}
    </div>
  );
}

export function Tag({ tone, children }: { tone?: 'ok' | 'warn' | 'bad' | 'busy' | 'folder'; children: ReactNode }) {
  return <span className={`tag ${tone ?? ''}`}>{children}</span>;
}

export function Note({ tone, children }: { tone?: 'warn' | 'bad'; children: ReactNode }) {
  return <div className={`note ${tone ?? ''}`} role={tone === 'bad' ? 'alert' : undefined}>{children}</div>;
}

export function More({ label = 'Advanced options', children, open }: { label?: string; children: ReactNode; open?: boolean }) {
  return (
    <details className="more" open={open}>
      <summary>{label}</summary>
      <div>{children}</div>
    </details>
  );
}

export function Check({ checked, onChange, label, hint, disabled }: { checked: boolean; onChange: (v: boolean) => void; label: ReactNode; hint?: ReactNode; disabled?: boolean }) {
  return (
    <label className="check">
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}{hint && <small>{hint}</small>}</span>
    </label>
  );
}

export function Field({ label, children, hint }: { label: string; children: ReactNode; hint?: string }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
      {hint && <small className="muted small">{hint}</small>}
    </label>
  );
}

export function Modal({ title, onClose, children, footer, narrow }: { title: string; onClose: () => void; children: ReactNode; footer?: ReactNode; narrow?: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  useEffect(() => {
    // Focus the dialog once when it opens (not on every re-render, which would steal focus from inputs).
    const before = document.activeElement as HTMLElement | null;
    ref.current?.focus();
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') closeRef.current(); };
    window.addEventListener('keydown', onKey);
    return () => { window.removeEventListener('keydown', onKey); before?.focus?.(); };
  }, []);
  return (
    <div className="modal-back" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={`modal ${narrow ? 'narrow' : ''}`} role="dialog" aria-modal="true" aria-label={title} tabIndex={-1} ref={ref}>
        <header><h2>{title}</h2><button className="icon-btn" onClick={onClose} aria-label="Close">×</button></header>
        <div className="body">{children}</div>
        {footer && <footer>{footer}</footer>}
      </div>
    </div>
  );
}

export interface ConfirmRequest { title: string; message: ReactNode; confirmLabel: string; danger?: boolean; onConfirm: () => void }

export function ConfirmDialog({ req, onClose }: { req: ConfirmRequest; onClose: () => void }) {
  return (
    <Modal title={req.title} onClose={onClose} narrow footer={<>
      <button className="btn" onClick={onClose}>Cancel</button>
      <button className={`btn ${req.danger ? 'danger' : 'primary'}`} onClick={() => { req.onConfirm(); onClose(); }}>{req.confirmLabel}</button>
    </>}>
      <div>{req.message}</div>
    </Modal>
  );
}

/** Multi-column picker: suggestions come from the detected headers; any name can be typed. */
export function ColumnPicker({ value, onChange, columns, placeholder = 'Type or pick a column…', single }: { value: string[]; onChange: (v: string[]) => void; columns: string[]; placeholder?: string; single?: boolean }) {
  const [text, setText] = useState('');
  const id = useId();
  const add = (v: string) => {
    const s = v.trim();
    if (!s) return;
    if (single) onChange([s]);
    else if (!value.some((x) => x.toLowerCase() === s.toLowerCase())) onChange([...value, s]);
    setText('');
  };
  return (
    <div className="chip-input">
      {value.map((v) => (
        <span className="chip" key={v}>{v}<button type="button" aria-label={`Remove ${v}`} onClick={() => onChange(value.filter((x) => x !== v))}>×</button></span>
      ))}
      {!(single && value.length) && (
        <input
          list={id}
          value={text}
          placeholder={value.length ? '' : placeholder}
          onChange={(e) => {
            const v = e.target.value;
            // Picking from the suggestion list fires without a typing inputType; typing "Unit" must not add "Unit" when "Units" was meant.
            const picked = !(e.nativeEvent as InputEvent).inputType || (e.nativeEvent as InputEvent).inputType === 'insertReplacementText';
            if (picked && columns.includes(v)) add(v); else setText(v);
          }}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ',') { e.preventDefault(); add(text); } if (e.key === 'Backspace' && !text && value.length) onChange(value.slice(0, -1)); }}
          onBlur={() => add(text)}
        />
      )}
      <datalist id={id}>{columns.filter((c) => !value.includes(c)).map((c) => <option key={c} value={c} />)}</datalist>
    </div>
  );
}

/** Single column chooser (select with free typing via datalist). */
export function ColumnSelect({ value, onChange, columns, placeholder = 'Column' }: { value: string; onChange: (v: string) => void; columns: string[]; placeholder?: string }) {
  const id = useId();
  return (
    <>
      <input type="text" list={id} value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />
      <datalist id={id}>{columns.map((c) => <option key={c} value={c} />)}</datalist>
    </>
  );
}

export function DataPreview({ headers, rows, total, rowNumbers, onPickRow, chosenRow }: { headers?: string[]; rows: (CellValue | string)[][]; total?: number; rowNumbers?: number[]; onPickRow?: (i: number) => void; chosenRow?: number }) {
  const width = Math.max(headers?.length ?? 0, ...rows.map((r) => r.length));
  return (
    <div>
      <div className="table-wrap">
        <table className="data">
          {headers && <thead><tr>{rowNumbers && <th>#</th>}{headers.map((h, i) => <th key={i} title={h}>{h}</th>)}</tr></thead>}
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} className={`${onPickRow ? 'pick' : ''} ${chosenRow === i ? 'chosen' : ''}`} onClick={onPickRow ? () => onPickRow(i) : undefined}>
                {rowNumbers && <td className="rn">{rowNumbers[i]}</td>}
                {Array.from({ length: width }, (_, c) => { const s = cellToString((r[c] ?? null) as CellValue); return <td key={c} title={s}>{s}</td>; })}
              </tr>
            ))}
            {!rows.length && <tr><td className="muted">No rows</td></tr>}
          </tbody>
        </table>
      </div>
      {total !== undefined && total > rows.length && <p className="muted small">Showing {rows.length} of {total.toLocaleString('en-IN')} rows.</p>}
    </div>
  );
}

export function Progress({ pct }: { pct: number }) {
  return <div className="progress" role="progressbar" aria-valuenow={pct} aria-valuemin={0} aria-valuemax={100}><div style={{ width: `${Math.max(0, Math.min(100, pct))}%` }} /></div>;
}
