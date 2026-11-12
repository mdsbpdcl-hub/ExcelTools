import { useState } from 'react';
import type { FileInfo } from '../../core/types';
import type { WorkflowConfig } from '../../core/workflow';
import type { PreviewResult, ProgressEvent } from '../../core/pipeline';
import { formatCount } from '../../core/util';
import { Check, DataPreview, More, Note, PageHead, Progress, Tag, Tiles } from '../components/common';

interface Props {
  preview: PreviewResult | null;
  previewing: boolean;
  progress: ProgressEvent | null;
  error: string | null;
  onPreview: () => void;
  onStop: () => void;
  acknowledged: boolean;
  setAcknowledged: (v: boolean) => void;
  config: WorkflowConfig;
  infos: FileInfo[];
}

const ISSUE_LABEL: Record<string, string> = {
  blank: 'Blank', 'invalid-date': 'Invalid date', 'unexpected-value': 'Unexpected value', 'spreadsheet-error': 'Spreadsheet error', validation: 'Rule failed', 'unsupported-formula': 'Formula not recalculated',
};

export function ReviewStep({ preview, previewing, progress, error, onPreview, onStop, acknowledged, setAcknowledged, config, infos }: Props) {
  const [open, setOpen] = useState<number>(0);
  const [sheetIdx, setSheetIdx] = useState(0);
  const formats = Object.entries(config.output.formats).filter(([, v]) => v).map(([k]) => k.toUpperCase());

  if (previewing || (!preview && !error)) {
    return (
      <>
        <PageHead title="Checking your workflow…" sub="The full batch is being processed in memory to show you exactly what will come out. Nothing is saved yet." />
        <div className="card">
          <Progress pct={progress?.pct ?? 5} />
          <p className="muted small">{progress?.message ?? 'Starting'}</p>
          {previewing ? <button className="btn small danger" onClick={onStop}>Stop</button> : <button className="btn primary" onClick={onPreview}>Prepare preview</button>}
        </div>
      </>
    );
  }
  if (error || !preview) {
    return (
      <>
        <PageHead title="The preview could not be prepared" />
        <Note tone="bad">{error}</Note>
        <button className="btn primary" onClick={onPreview}>Try again</button>
      </>
    );
  }
  const out = preview.outputs[open];
  const sheet = out?.sheets[Math.min(sheetIdx, (out?.sheets.length ?? 1) - 1)];
  const profileWarnings = preview.profiles.reduce((s, p) => s + p.warnings.length, 0);

  return (
    <>
      <PageHead title="Check before processing" sub={`This is what will be produced, in ${formats.join(', ') || 'no format'}. Nothing has been saved yet.`} />
      <Tiles items={[
        { label: 'Files to process', value: infos.filter((i) => i.status === 'ready' && !preview.errors.some((e) => e.fileId === i.id)).length },
        { label: 'Rows read', value: formatCount(preview.stats['Rows read'] ?? 0) },
        { label: 'Outputs', value: preview.outputs.length },
        { label: 'Issues found', value: formatCount(preview.issueCount), tone: preview.issueCount ? 'warn' : 'ok' },
      ]} />

      {preview.destructive.length > 0 && (
        <Note tone="warn">
          <b>This run removes data from the output.</b> Your original files are not changed.
          <ul>{preview.destructive.map((d) => <li key={d}>{d}</li>)}</ul>
          <Check checked={acknowledged} onChange={setAcknowledged} label="I understand — continue" />
        </Note>
      )}
      {preview.errors.length > 0 && (
        <Note tone="bad">
          <b>{preview.errors.length} problem(s) — these go to error/ and the rest continues:</b>
          <ul>{preview.errors.map((e, i) => <li key={i}><b>{e.relPath || e.name}</b>: {e.reason}{e.suggestion ? ` — ${e.suggestion}` : ''}</li>)}</ul>
        </Note>
      )}

      <div className="card">
        <h2>Outputs</h2>
        <p className="sub">Click an output to see its first rows.</p>
        <div className="list">
          {preview.outputs.map((o, i) => (
            <div className="list-item" key={i} style={{ cursor: 'pointer', background: i === open ? 'var(--accent-tint)' : undefined }} onClick={() => { setOpen(i); setSheetIdx(0); }}>
              <Tag tone="folder">{o.folder}/</Tag>
              <span className="name">{o.name}</span>
              <span className="meta">{o.sheets.length > 1 ? `${o.sheets.length} sheets · ` : ''}{formatCount(o.sheets.reduce((s, x) => s + x.totalRows, 0))} rows · {o.sheets[0]?.headers.length ?? 0} columns</span>
            </div>
          ))}
          {!preview.outputs.length && <div className="list-item muted">No outputs — check the errors above.</div>}
        </div>
        {out && sheet && (
          <div style={{ marginTop: 12 }}>
            {out.note && <p className="muted small">{out.note}</p>}
            {out.sheets.length > 1 && (
              <div className="chips" style={{ marginBottom: 8 }}>
                {out.sheets.map((s, i) => <button key={i} className={`chip ${i === sheetIdx ? 'on' : ''}`} onClick={() => setSheetIdx(i)}>{s.name}</button>)}
              </div>
            )}
            <DataPreview headers={sheet.headers} rows={sheet.rows} total={sheet.totalRows} />
          </div>
        )}
      </div>

      <div className="card">
        <h2>What will be done</h2>
        {Object.keys(preview.stats).length ? (
          <div className="grid3">{Object.entries(preview.stats).map(([k, v]) => <div key={k}><span className="muted small">{k}</span><div><b>{formatCount(v)}</b></div></div>)}</div>
        ) : <p className="muted">Files are converted to the chosen formats.</p>}
        <Note>{preview.formulaNote}</Note>
      </div>

      <div className="card">
        <h2>Data profile</h2>
        <p className="sub">{preview.profiles.length} table(s) · {profileWarnings} warning(s)</p>
        {preview.profiles.map((p, i) => (
          <More key={i} label={`${p.file}${p.sheet && p.sheet !== p.file ? ` › ${p.sheet}` : ''} — ${formatCount(p.rows)} rows, ${p.columns} columns${p.warnings.length ? `, ${p.warnings.length} warning(s)` : ''}`}>
            {p.warnings.length > 0 && <Note tone="warn"><ul>{p.warnings.map((w) => <li key={w}>{w}</li>)}</ul></Note>}
            <div className="table-wrap" style={{ maxHeight: 280 }}>
              <table className="data">
                <thead><tr><th>Column</th><th>Type</th><th>Blanks</th><th>Unique</th><th>Min</th><th>Max</th><th>Errors</th><th>Examples</th></tr></thead>
                <tbody>{p.columnProfiles.map((c) => (
                  <tr key={c.name}><td>{c.name}</td><td>{c.type}</td><td>{formatCount(c.blanks)}</td><td>{formatCount(c.unique)}{c.uniqueCapped ? '+' : ''}</td><td>{c.min ?? ''}</td><td>{c.max ?? ''}</td><td>{c.errors || ''}</td><td>{c.sample.join(' · ')}</td></tr>
                ))}</tbody>
              </table>
            </div>
          </More>
        ))}
      </div>

      {preview.issueCount > 0 && (
        <div className="card">
          <h2>Issues ({formatCount(preview.issueCount)})</h2>
          <p className="sub">All issues will be listed in review/Issues report{preview.issueCount > preview.issues.length ? `; the first ${preview.issues.length} are shown here` : ''}.</p>
          <DataPreview headers={['Issue', 'File', 'Sheet', 'Row', 'Column', 'Value', 'Details']} rows={preview.issues.map((i) => [ISSUE_LABEL[i.kind], i.file, i.sheet, i.row, i.column, i.value, i.message])} />
        </div>
      )}
      {preview.warnings.length > 0 && (
        <div className="card">
          <h2>Warnings ({preview.warnings.length})</h2>
          <ul className="small">{preview.warnings.slice(0, 100).map((w, i) => <li key={i}>{w}</li>)}</ul>
          {preview.warnings.length > 100 && <p className="muted small">…and {preview.warnings.length - 100} more in the run report.</p>}
        </div>
      )}
      <button className="btn" onClick={onPreview}>Refresh preview</button>
    </>
  );
}
