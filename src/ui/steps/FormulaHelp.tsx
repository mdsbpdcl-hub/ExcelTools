import { useEffect, useState } from 'react';
import { checkFormula, FUNCTION_REFERENCE, supportedFunctionNames, type FormulaCheck } from '../../core/formulas';
import { Modal, Note, Tag } from '../components/common';

export function FormulaHelp({ onClose }: { onClose: () => void }) {
  const [supported, setSupported] = useState<Set<string> | null>(null);
  const [cat, setCat] = useState('All');
  const [q, setQ] = useState('');
  const [formula, setFormula] = useState('');
  const [check, setCheck] = useState<FormulaCheck | null>(null);
  useEffect(() => { void supportedFunctionNames().then(setSupported); }, []);
  const cats = ['All', ...new Set(FUNCTION_REFERENCE.map((f) => f.category))];
  const list = FUNCTION_REFERENCE.filter((f) => (cat === 'All' || f.category === cat) && (!q || `${f.name} ${f.description}`.toLowerCase().includes(q.toLowerCase())));
  return (
    <Modal title="Formula help" onClose={onClose}>
      <Note>
        <b>Preserved</b> means a formula is copied into the output and Excel recalculates it when you open the file.{' '}
        <b>Recalculated</b> means this app worked out new results (only when “Recalculate formulas” is on, and only for functions marked “Can recalculate”).
        The run report always says which happened. {supported && <>The engine here supports {supported.size} Excel functions.</>}
      </Note>
      <div className="field-label">Check a formula</div>
      <div className="row">
        <input type="text" className="mono grow" placeholder="=XLOOKUP(A2, Master!A:A, Master!C:C)" value={formula} onChange={(e) => setFormula(e.target.value)} onKeyDown={async (e) => { if (e.key === 'Enter') setCheck(await checkFormula(formula)); }} style={{ flex: 1 }} />
        <button className="btn" onClick={async () => setCheck(await checkFormula(formula))}>Check</button>
      </div>
      {check && <Note tone={!check.valid ? 'bad' : check.functions.some((f) => !f.supported) ? 'warn' : undefined}>
        {check.message}
        {check.functions.length > 0 && <div style={{ marginTop: 6 }}>{check.functions.map((f) => <span key={f.name} style={{ marginRight: 6 }}><Tag tone={f.supported ? 'ok' : 'warn'}>{f.name}</Tag></span>)}</div>}
      </Note>}
      <div className="row" style={{ margin: '14px 0 8px' }}>
        <div className="chips">{cats.map((c) => <button key={c} className={`chip ${c === cat ? 'on' : ''}`} onClick={() => setCat(c)}>{c}</button>)}</div>
        <input type="text" placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} style={{ maxWidth: 200, marginLeft: 'auto' }} />
      </div>
      <div className="list">
        {list.map((f) => (
          <div className="list-item" key={f.name} style={{ alignItems: 'flex-start' }}>
            <div className="name" style={{ whiteSpace: 'normal' }}>
              <b className="mono">{f.name}</b> <span className="muted small">· {f.category}</span>
              <div className="small">{f.description}</div>
              <div className="mono small muted">{f.syntax}</div>
              <div className="mono small">{f.example}</div>
            </div>
            {supported ? (supported.has(f.name) ? <Tag tone="ok">Can recalculate</Tag> : <Tag tone="warn">Preserved only</Tag>) : <Tag>…</Tag>}
          </div>
        ))}
      </div>
    </Modal>
  );
}
