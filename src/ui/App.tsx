import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { FileInfo, SourceFile } from '../core/types';
import { defaultWorkflow, type FileReadOverride, type WorkflowConfig } from '../core/workflow';
import type { PreviewResult, ProgressEvent, RunResult } from '../core/pipeline';
import { isIgnorableName, kindFromName } from '../core/parse';
import { uid } from '../core/util';
import { CancelledError, TaskFailed, WorkerClient } from './workerClient';
import { ConfirmDialog, Stepper, type ConfirmRequest } from './components/common';
import { FilesStep } from './steps/FilesStep';
import { FileSettings } from './steps/FileSettings';
import { TasksStep } from './steps/TasksStep';
import { SettingsStep } from './steps/SettingsStep';
import { ReviewStep } from './steps/ReviewStep';
import { ProcessStep, type FileRunState } from './steps/ProcessStep';
import { ResultsStep } from './steps/ResultsStep';
import { FormulaHelp } from './steps/FormulaHelp';

const STEPS = ['Files', 'Tasks', 'Settings', 'Review', 'Process', 'Results'];

export interface IncomingFile { file: File; relPath: string }

export function App() {
  const client = useRef<WorkerClient | null>(null);
  if (!client.current) client.current = new WorkerClient();
  const worker = client.current;

  const [step, setStep] = useState(0);
  const [maxReached, setMaxReached] = useState(0);
  const [sources, setSources] = useState<SourceFile[]>([]);
  const [infos, setInfos] = useState<FileInfo[]>([]);
  const [checking, setChecking] = useState<Set<string>>(new Set());
  const [config, setConfig] = useState<WorkflowConfig>(defaultWorkflow);
  const [settingsFor, setSettingsFor] = useState<string | null>(null);
  const [preview, setPreview] = useState<PreviewResult | null>(null);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [previewing, setPreviewing] = useState(false);
  const [previewProgress, setPreviewProgress] = useState<ProgressEvent | null>(null);
  const [acknowledged, setAcknowledged] = useState(false);
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<ProgressEvent | null>(null);
  const [fileStates, setFileStates] = useState<Record<string, FileRunState>>({});
  const [log, setLog] = useState<string[]>([]);
  const [runError, setRunError] = useState<{ message: string; suggestion?: string } | null>(null);
  const [result, setResult] = useState<RunResult | null>(null);
  const [confirm, setConfirm] = useState<ConfirmRequest | null>(null);
  const [showHelp, setShowHelp] = useState(false);

  /** Any change to files or settings makes the preview and the previous run out of date. */
  const invalidate = useCallback(() => {
    setPreview(null);
    setPreviewError(null);
    setAcknowledged(false);
    setResult(null);
    setRunError(null);
    setFileStates({});
    setLog([]);
  }, []);

  const update = useCallback((fn: (draft: WorkflowConfig) => void) => {
    setConfig((c) => { const d = structuredClone(c); fn(d); return d; });
    invalidate();
  }, []);

  const go = (i: number) => {
    setStep(i);
    setMaxReached((m) => Math.max(m, i));
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  // ---------- files ----------
  const inspect = useCallback(async (list: SourceFile[], overrides: Record<string, FileReadOverride>) => {
    if (!list.length) return;
    setChecking((s) => new Set([...s, ...list.map((f) => f.id)]));
    try {
      const res = await worker.request<{ type: 'inspected'; reqId: number; infos: FileInfo[] }>({ type: 'inspect', files: list, overrides });
      setInfos((cur) => {
        const map = new Map(cur.map((i) => [i.id, i]));
        for (const i of res.infos) map.set(i.id, i);
        return [...map.values()];
      });
    } catch (e) {
      if (!(e instanceof CancelledError)) {
        const msg = e instanceof Error ? e.message : String(e);
        setInfos((cur) => [...cur.filter((i) => !list.some((f) => f.id === i.id)), ...list.map((f) => ({ id: f.id, name: f.name, relPath: f.relPath, size: f.size, kind: f.kind, status: 'error' as const, reason: msg, sheets: [], warnings: [] }))]);
      }
    } finally {
      setChecking((s) => { const n = new Set(s); list.forEach((f) => n.delete(f.id)); return n; });
    }
  }, [worker]);

  const addFiles = (incoming: IncomingFile[]) => {
    const known = new Set(sources.map((s) => `${s.relPath}|${s.size}`));
    const fresh: SourceFile[] = [];
    let ignored = 0;
    for (const { file, relPath } of incoming) {
      if (isIgnorableName(file.name)) { ignored++; continue; }
      const key = `${relPath}|${file.size}`;
      if (known.has(key)) continue;
      known.add(key);
      fresh.push({ id: uid('f'), name: file.name, relPath, size: file.size, kind: kindFromName(file.name), file });
    }
    if (!fresh.length && !ignored) return;
    setSources((s) => [...s, ...fresh]);
    invalidate();
    void inspect(fresh, {});
  };

  const removeFile = (id: string) => {
    setSources((s) => s.filter((f) => f.id !== id));
    setInfos((s) => s.filter((f) => f.id !== id));
    update((d) => { delete d.read.files[id]; });
  };

  const clearAll = () => setConfirm({
    title: 'Remove all files?', message: 'The list will be cleared. Your files on this computer are not changed.', confirmLabel: 'Remove all', danger: true,
    onConfirm: () => { setSources([]); setInfos([]); update((d) => { d.read.files = {}; }); },
  });

  const saveFileSettings = (id: string, ov: FileReadOverride) => {
    update((d) => { d.read.files[id] = ov; });
    const src = sources.find((s) => s.id === id);
    if (src) void inspect([src], { [id]: ov });
    setSettingsFor(null);
  };

  /** Column names found in the selected sheets (with the chosen header rows), for the column pickers. */
  const columns = useMemo(() => {
    const out: string[] = [];
    for (const info of infos) {
      if (info.status !== 'ready') continue;
      const ov = config.read.files[info.id];
      for (const s of info.sheets) {
        if (ov?.sheets ? !ov.sheets.includes(s.name) : s.hidden && !config.read.includeHidden) continue;
        for (const h of s.headers) if (!out.includes(h)) out.push(h);
      }
    }
    for (const r of config.columns.derived) if (r.name && !out.includes(r.name)) out.push(r.name);
    for (const r of config.columns.rename) if (r.to && !out.includes(r.to)) out.push(r.to);
    for (const r of config.columns.combine) if (r.into && !out.includes(r.into)) out.push(r.into);
    for (const r of config.columns.split) for (const n of r.into) if (n && !out.includes(n)) out.push(n);
    if (config.combine.mode === 'merge') { for (const n of ['Source file', 'Source sheet', 'Source row']) if (!out.includes(n)) out.push(n); }
    return out;
  }, [infos, config.read, config.columns, config.combine.mode]);

  // ---------- preview & run ----------
  const runPreview = async () => {
    setPreviewing(true);
    setPreviewError(null);
    setPreviewProgress(null);
    try {
      const res = await worker.request<{ type: 'previewed'; reqId: number; result: PreviewResult }>({ type: 'preview', files: sources, infos, config }, setPreviewProgress);
      setPreview(res.result);
    } catch (e) {
      if (!(e instanceof CancelledError)) setPreviewError(e instanceof Error ? e.message + (e instanceof TaskFailed && e.suggestion ? ` ${e.suggestion}` : '') : String(e));
    } finally {
      setPreviewing(false);
    }
  };

  const stopPreview = () => { worker.cancel(); setPreviewError('Preview stopped. Change the settings or try again.'); };

  useEffect(() => {
    if (step === 3 && !preview && !previewing && !previewError) void runPreview();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [step, preview, previewing, previewError]);

  const startRun = async () => {
    setRunning(true);
    setRunError(null);
    setResult(null);
    setLog([]);
    setProgress({ phase: 'read', pct: 0, message: 'Starting' });
    const states: Record<string, FileRunState> = {};
    for (const i of infos) states[i.id] = i.status === 'ready' ? 'queued' : i.status === 'skipped' ? 'skipped' : 'error';
    setFileStates(states);
    const stamp = () => new Date().toLocaleTimeString('en-IN', { hour12: false });
    try {
      const res = await worker.request<{ type: 'done'; reqId: number; result: RunResult }>({ type: 'run', files: sources, infos, config }, (e) => {
        setProgress(e);
        if (e.fileId && e.fileState) setFileStates((s) => ({ ...s, [e.fileId!]: e.fileState! }));
        setLog((l) => (l[l.length - 1]?.endsWith(e.message) ? l : [...l.slice(-300), `${stamp()}  ${e.message}`]));
      });
      setResult(res.result);
      setLog((l) => [...l, `${stamp()}  Finished: ${res.result.files.length} file(s) in the package.`]);
      go(5);
    } catch (e) {
      if (e instanceof CancelledError) {
        setLog((l) => [...l, `${stamp()}  Cancelled. Nothing was saved.`]);
        setRunError({ message: 'Processing was cancelled. Nothing was saved; your files were not changed.' });
      } else {
        setRunError({ message: e instanceof Error ? e.message : String(e), suggestion: e instanceof TaskFailed ? e.suggestion : undefined });
      }
    } finally {
      setRunning(false);
    }
  };

  const cancelRun = () => setConfirm({
    title: 'Stop processing?', message: 'The current run will stop now. Nothing is saved from a stopped run, and your original files are never changed.', confirmLabel: 'Stop', danger: true,
    onConfirm: () => worker.cancel(),
  });

  const newBatch = () => setConfirm({
    title: 'Start a new batch?', message: 'This clears the file list and the results on screen. Download anything you need first. Your settings are kept.', confirmLabel: 'Start new batch',
    onConfirm: () => { setSources([]); setInfos([]); setPreview(null); setResult(null); setFileStates({}); setLog([]); setMaxReached(0); update((d) => { d.read.files = {}; }); go(0); },
  });

  // ---------- navigation rules ----------
  const ready = infos.filter((i) => i.status === 'ready');
  const formatsChosen = Object.values(config.output.formats).some(Boolean);
  const needsAck = !!preview?.destructive.length && !acknowledged;
  const canNext = [
    ready.length > 0 && checking.size === 0,
    formatsChosen,
    true,
    !!preview && !previewing && !needsAck,
    false,
    false,
  ][step];
  const nextHint = [
    checking.size ? 'Checking files…' : !ready.length ? 'Add at least one file that can be read.' : `${ready.length} file(s) ready`,
    formatsChosen ? '' : 'Choose at least one output format.',
    '',
    previewing ? 'Preparing preview…' : needsAck ? 'Please confirm the data removal notice above.' : '',
    running ? 'Processing…' : '',
    '',
  ][step];

  return (
    <>
      <header className="app-header">
        <div className="inner">
          <div className="logo" aria-hidden="true">SD</div>
          <div>
            <div className="app-name">Sheet Desk</div>
            <div className="app-sub">Batch processing for Excel and CSV files</div>
          </div>
          <div className="header-right">
            <button className="btn small" onClick={() => setShowHelp(true)}>ƒx Formula help</button>
            <span title="Files are read and processed inside this browser tab. Nothing is uploaded."><span className="privacy-dot" /><span className="privacy-text">Files stay on this computer</span></span>
          </div>
        </div>
      </header>
      <main>
        <Stepper steps={STEPS} current={step} maxReached={running ? step : result ? 5 : Math.min(maxReached, 3)} onGo={(i) => !running && go(i)} />
        {step === 0 && <FilesStep sources={sources} infos={infos} checking={checking} config={config} update={update} onAdd={addFiles} onRemove={removeFile} onClear={clearAll} onSettings={setSettingsFor} />}
        {step === 1 && <TasksStep config={config} update={update} setConfig={(c) => { setConfig(c); invalidate(); }} fileCount={ready.length} />}
        {step === 2 && <SettingsStep config={config} update={update} columns={columns} infos={infos} />}
        {step === 3 && <ReviewStep preview={preview} previewing={previewing} progress={previewProgress} error={previewError} onPreview={runPreview} onStop={stopPreview} acknowledged={acknowledged} setAcknowledged={setAcknowledged} config={config} infos={infos} />}
        {step === 4 && <ProcessStep infos={infos} running={running} progress={progress} states={fileStates} log={log} error={runError} onStart={startRun} onCancel={cancelRun} done={!!result} />}
        {step === 5 && result && <ResultsStep result={result} onNewBatch={newBatch} onAdjust={() => go(1)} />}
      </main>
      <div className="bottom-bar">
        <div className="inner">
          <button className="btn" disabled={step === 0 || running} onClick={() => go(step - 1)}>← Back</button>
          <div className="hint">{nextHint}</div>
          {step < 4 && (
            <button className="btn primary" disabled={!canNext} onClick={() => go(step + 1)}>
              {step === 3 ? 'Continue to processing →' : 'Next →'}
            </button>
          )}
          {step === 4 && !running && !result && <button className="btn primary" onClick={startRun}>Start processing</button>}
          {step === 4 && running && <button className="btn danger" onClick={cancelRun}>Stop</button>}
          {step === 4 && result && !running && <button className="btn primary" onClick={() => go(5)}>See results →</button>}
          {step === 5 && <button className="btn primary" onClick={newBatch}>New batch</button>}
        </div>
      </div>
      {settingsFor && (() => {
        const info = infos.find((i) => i.id === settingsFor);
        return info ? <FileSettings info={info} override={config.read.files[info.id] ?? {}} includeHidden={config.read.includeHidden} onClose={() => setSettingsFor(null)} onSave={(ov) => saveFileSettings(info.id, ov)} /> : null;
      })()}
      {confirm && <ConfirmDialog req={confirm} onClose={() => setConfirm(null)} />}
      {showHelp && <FormulaHelp onClose={() => setShowHelp(false)} />}
    </>
  );
}
