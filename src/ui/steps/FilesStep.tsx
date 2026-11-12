import { useRef, useState } from 'react';
import type { FileInfo, SourceFile } from '../../core/types';
import type { WorkflowConfig } from '../../core/workflow';
import { LIMITS } from '../../core/limits';
import { formatBytes, formatCount } from '../../core/util';
import { Check, More, Note, Tag, Tiles } from '../components/common';
import type { IncomingFile } from '../App';

interface Props {
  sources: SourceFile[];
  infos: FileInfo[];
  checking: Set<string>;
  config: WorkflowConfig;
  update: (fn: (d: WorkflowConfig) => void) => void;
  onAdd: (files: IncomingFile[]) => void;
  onRemove: (id: string) => void;
  onClear: () => void;
  onSettings: (id: string) => void;
}

/** Reads dropped folders recursively (Chrome, Edge, Firefox, Safari all support webkitGetAsEntry). */
async function filesFromDrop(items: DataTransferItemList): Promise<IncomingFile[]> {
  const out: IncomingFile[] = [];
  const walk = async (entry: FileSystemEntry, path: string): Promise<void> => {
    if (entry.isFile) {
      const file = await new Promise<File>((res, rej) => (entry as FileSystemFileEntry).file(res, rej));
      out.push({ file, relPath: path + file.name });
    } else if (entry.isDirectory) {
      const reader = (entry as FileSystemDirectoryEntry).createReader();
      for (;;) {
        const batch = await new Promise<FileSystemEntry[]>((res, rej) => reader.readEntries(res, rej));
        if (!batch.length) break;
        for (const e of batch) await walk(e, `${path}${entry.name}/`);
      }
    }
  };
  const entries = [...items].map((i) => i.webkitGetAsEntry?.()).filter(Boolean) as FileSystemEntry[];
  for (const e of entries) await walk(e, '');
  return out;
}

export function FilesStep({ sources, infos, checking, config, update, onAdd, onRemove, onClear, onSettings }: Props) {
  const fileInput = useRef<HTMLInputElement>(null);
  const folderInput = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  const ready = infos.filter((i) => i.status === 'ready');
  const problems = infos.filter((i) => i.status !== 'ready');
  const total = sources.reduce((s, f) => s + f.size, 0);
  const rows = ready.reduce((s, i) => s + i.sheets.reduce((a, sh) => a + Math.max(0, sh.rowCount - 1), 0), 0);

  const fromInput = (list: FileList | null) => {
    if (!list) return;
    onAdd([...list].map((file) => ({ file, relPath: (file as File & { webkitRelativePath?: string }).webkitRelativePath || file.name })));
  };

  return (
    <>
      <div className="page-head">
        <h1>Which files do you want to process?</h1>
        <p>Add Excel or CSV files, or a whole folder (sub-folders are included). Your original files are never changed.</p>
      </div>
      {sources.length > 0 && (
        <Tiles items={[
          { label: 'Files added', value: sources.length },
          { label: 'Ready', value: ready.length, tone: 'ok' },
          { label: 'Skipped or problems', value: problems.length, tone: problems.length ? 'warn' : undefined },
          { label: 'Total size', value: formatBytes(total) },
        ]} />
      )}
      <div
        className={`drop ${over ? 'over' : ''}`}
        onDragOver={(e) => { e.preventDefault(); setOver(true); }}
        onDragLeave={() => setOver(false)}
        onDrop={async (e) => { e.preventDefault(); setOver(false); onAdd(await filesFromDrop(e.dataTransfer.items)); }}
      >
        <b>Drop files or folders here</b>
        <p>.xlsx · .xlsm · .xls · .xlsb · .csv · .tsv · .txt</p>
        <div className="row" style={{ justifyContent: 'center' }}>
          <button className="btn primary" onClick={() => fileInput.current?.click()}>Choose files</button>
          <button className="btn" onClick={() => folderInput.current?.click()}>Choose a folder</button>
        </div>
        <input ref={fileInput} type="file" multiple hidden accept=".xlsx,.xlsm,.xls,.xlsb,.csv,.tsv,.txt" onChange={(e) => { fromInput(e.target.files); e.target.value = ''; }} />
        <input ref={folderInput} type="file" hidden {...{ webkitdirectory: '', directory: '' }} onChange={(e) => { fromInput(e.target.files); e.target.value = ''; }} />
      </div>
      <p className="muted small">
        Limits: {LIMITS.maxFileBytes / 1048576} MB per file, {LIMITS.maxBatchBytes / 1048576} MB per batch, {formatCount(LIMITS.maxRowsPerTable)} rows per sheet. Files above a limit are stopped with a message — data is never cut off silently.
        Browsers cannot open a folder on your computer directly; choosing a folder reads a copy of its files into this tab.
      </p>

      {sources.length > 0 && (
        <div className="card">
          <div className="card-head">
            <h2>Files</h2>
            <span className="muted small">{formatCount(rows)} rows found in ready files</span>
            <span className="spacer" />
            <button className="btn small danger" onClick={onClear}>Remove all</button>
          </div>
          <div className="list">
            {sources.map((s) => {
              const info = infos.find((i) => i.id === s.id);
              const busy = checking.has(s.id) || !info;
              const ov = config.read.files[s.id];
              const sheetsUsed = info?.sheets.filter((sh) => (ov?.sheets ? ov.sheets.includes(sh.name) : !sh.hidden || config.read.includeHidden)).length ?? 0;
              return (
                <div className="list-item" key={s.id}>
                  <div className="name">
                    <div title={s.relPath}>{s.name}</div>
                    {s.relPath !== s.name && <div className="path">{s.relPath}</div>}
                    {info?.reason && <div className="reason">{info.reason}{info.suggestion ? ` — ${info.suggestion}` : ''}</div>}
                    {info?.warnings.map((w) => <div className="reason" key={w}>⚠ {w}</div>)}
                  </div>
                  {info?.status === 'ready' && <span className="meta">{info.kind === 'csv' || info.kind === 'tsv' || info.kind === 'txt' ? `${info.delimiter === '\t' ? 'Tab' : info.delimiter}-separated · ${info.encoding}` : `${sheetsUsed} of ${info.sheets.length} sheet(s)`}</span>}
                  <span className="meta">{formatBytes(s.size)}</span>
                  {busy ? <Tag tone="busy">Checking</Tag> : info!.status === 'ready' ? <Tag tone="ok">Ready</Tag> : info!.status === 'skipped' ? <Tag>Skipped</Tag> : <Tag tone="bad">Problem</Tag>}
                  <button className="icon-btn" title="Sheets, header row and text settings" aria-label={`Settings for ${s.name}`} disabled={info?.status !== 'ready'} onClick={() => onSettings(s.id)}>⚙</button>
                  <button className="icon-btn" title="Remove from list" aria-label={`Remove ${s.name}`} onClick={() => onRemove(s.id)}>×</button>
                </div>
              );
            })}
          </div>
          {problems.some((p) => p.status === 'error') && <Note tone="warn">Files with problems are not processed. They are listed in the run report with the reason, and (if you keep the option on) copied into the <b>error/</b> folder.</Note>}
          <More label="Reading options">
            <Check checked={config.read.includeHidden} onChange={(v) => update((d) => { d.read.includeHidden = v; })} label="Include hidden sheets" />
            <Check checked={config.read.mergedCells === 'fill'} onChange={(v) => update((d) => { d.read.mergedCells = v ? 'fill' : 'first'; })} label="Fill merged cells with their value" hint="Otherwise only the top-left cell of a merged area keeps the value." />
            <Check checked={config.read.skipBlankRows} onChange={(v) => update((d) => { d.read.skipBlankRows = v; })} label="Remove completely blank rows" hint="Removing rows changes cell positions, so formulas in those sheets are saved as values." />
            <Check checked={config.read.convertTextNumbers} onChange={(v) => update((d) => { d.read.convertTextNumbers = v; })} label="Treat plain numbers in CSV/text files as numbers" hint="IDs with leading zeros or more than 15 digits always stay as text." />
          </More>
        </div>
      )}
    </>
  );
}

