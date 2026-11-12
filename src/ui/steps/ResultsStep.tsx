import { useState } from 'react';
import type { RunResult } from '../../core/pipeline';
import { formatBytes, formatCount } from '../../core/util';
import { More, Note, PageHead, Tag, Tiles } from '../components/common';

function download(bytes: Uint8Array, name: string) {
  const blob = new Blob([bytes as BlobPart]);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name.split('/').pop() || name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

type DirHandle = { name: string; getDirectoryHandle: (n: string, o?: { create?: boolean }) => Promise<DirHandle>; getFileHandle: (n: string, o?: { create?: boolean }) => Promise<{ createWritable: () => Promise<{ write: (b: BlobPart) => Promise<void>; close: () => Promise<void> }> }> };

const canPickFolder = typeof window !== 'undefined' && 'showDirectoryPicker' in window;

/**
 * Chrome/Edge only: writes the package into a NEW sub-folder of a folder the user picks, so nothing
 * outside that folder can be touched and existing files are never overwritten.
 */
async function saveToFolder(result: RunResult): Promise<string> {
  const root = await (window as unknown as { showDirectoryPicker: (o: object) => Promise<DirHandle> }).showDirectoryPicker({ mode: 'readwrite', id: 'sheet-desk-out' });
  const base = result.zipName.replace(/\.zip$/, '');
  let dirName = base;
  for (let i = 2; ; i++) {
    try { await root.getDirectoryHandle(dirName); dirName = `${base} (${i})`; } catch { break; }
  }
  const outDir = await root.getDirectoryHandle(dirName, { create: true });
  for (const f of result.files) {
    const parts = f.path.split('/').filter((p) => p && p !== '.' && p !== '..');
    let dir = outDir;
    for (const p of parts.slice(0, -1)) dir = await dir.getDirectoryHandle(p, { create: true });
    const fh = await dir.getFileHandle(parts[parts.length - 1], { create: true });
    const w = await fh.createWritable();
    await w.write(f.bytes as BlobPart);
    await w.close();
  }
  return `${root.name}/${dirName}`;
}

export function ResultsStep({ result, onNewBatch, onAdjust }: { result: RunResult; onNewBatch: () => void; onAdjust: () => void }) {
  const s = result.summary;
  const [saved, setSaved] = useState<string | null>(null);
  const [saveErr, setSaveErr] = useState<string | null>(null);
  const groups = (['success', 'review', 'error', 'root'] as const).map((g) => ({ g, files: result.files.filter((f) => f.folder === g) })).filter((x) => x.files.length);
  const fileWarnings = s.files.flatMap((f) => f.warnings.map((w) => `${f.path || f.name}: ${w}`));
  return (
    <>
      <PageHead title={s.totals.failed ? 'Finished, with some problems' : 'All done'} sub={`Processed in ${s.durationSeconds} s. ${formatCount(s.totals.rowsRead)} rows read, ${formatCount(s.totals.rowsWritten)} rows written.`} />
      <Tiles items={[
        { label: 'Succeeded', value: s.totals.succeeded, tone: 'ok' },
        { label: 'Need review', value: s.totals.needsReview, tone: s.totals.needsReview ? 'warn' : undefined },
        { label: 'Failed', value: s.totals.failed, tone: s.totals.failed ? 'bad' : undefined },
        { label: 'Skipped', value: s.totals.skipped },
      ]} />
      <div className="card">
        <div className="row">
          <button className="btn primary" onClick={() => download(result.zip, result.zipName)}>⬇ Download ZIP ({formatBytes(result.zip.length)})</button>
          {canPickFolder && <button className="btn" onClick={async () => { setSaveErr(null); try { setSaved(await saveToFolder(result)); } catch (e) { if ((e as Error).name !== 'AbortError') setSaveErr((e as Error).message); } }}>Save into a folder…</button>}
          <span className="grow" />
          <button className="btn" onClick={onAdjust}>Change settings and run again</button>
        </div>
        {saved && <Note>Saved into a new folder: <b>{saved}</b></Note>}
        {saveErr && <Note tone="bad">Could not save to the folder: {saveErr}</Note>}
        <p className="muted small">The ZIP holds success/ (finished), review/ (needs a look, plus reports) and error/ (could not be processed), with run-summary.json and run-report.txt.</p>
      </div>
      <Note>{s.formulas.statement}</Note>
      {s.errors.length > 0 && (
        <Note tone="bad">
          <b>Problems</b>
          <ul>{s.errors.map((e, i) => <li key={i}><b>{e.relPath || e.name}</b>: {e.reason}{e.suggestion && <> — <i>{e.suggestion}</i></>}</li>)}</ul>
        </Note>
      )}
      <div className="card">
        <h2>Files in the package</h2>
        {groups.map(({ g, files }) => (
          <div key={g} style={{ marginTop: 10 }}>
            <div className="field-label">{g === 'root' ? 'Reports' : `${g}/`}</div>
            <div className="list">
              {files.map((f) => (
                <div className="list-item" key={f.path}>
                  <Tag tone={g === 'success' ? 'ok' : g === 'review' ? 'warn' : g === 'error' ? 'bad' : undefined}>{g === 'root' ? 'report' : g}</Tag>
                  <span className="name" title={f.path}>{f.path.replace(/^(success|review|error)\//, '')}</span>
                  <span className="meta">{formatBytes(f.size)}</span>
                  <button className="icon-btn" aria-label={`Download ${f.path}`} title="Download this file" onClick={() => download(f.bytes, f.path)}>⬇</button>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
      {(s.warnings.length > 0 || fileWarnings.length > 0) && (
        <div className="card">
          <More label={`Warnings (${s.warnings.length + fileWarnings.length})`}>
            <ul className="small">{[...s.warnings, ...fileWarnings].map((w, i) => <li key={i}>{w}</li>)}</ul>
          </More>
        </div>
      )}
      <div className="row"><button className="btn" onClick={onNewBatch}>Start a new batch</button></div>
    </>
  );
}
