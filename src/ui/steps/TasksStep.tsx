import { useRef, useState } from 'react';
import { upgradeWorkflow, type WorkflowConfig } from '../../core/workflow';
import { Note, PageHead } from '../components/common';

interface Props {
  config: WorkflowConfig;
  update: (fn: (d: WorkflowConfig) => void) => void;
  setConfig: (c: WorkflowConfig) => void;
  fileCount: number;
}

type ToggleKey = 'errors' | 'clean' | 'findReplace' | 'columns' | 'filter' | 'dedupe' | 'validate' | 'split' | 'pivot' | 'format';

const GROUPS: { title: string; items: { key: ToggleKey; label: string; hint: string }[] }[] = [
  {
    title: 'Clean and check',
    items: [
      { key: 'clean', label: 'Clean up text, dates and numbers', hint: 'Trim spaces, fix case, remove hidden characters, standardise dates and numbers' },
      { key: 'findReplace', label: 'Find and replace', hint: 'Exact, partial or pattern (regex) replacements across files' },
      { key: 'errors', label: 'Find data problems', hint: '#N/A and other errors, blank cells, invalid dates, unexpected values' },
      { key: 'validate', label: 'Check rules', hint: 'Required fields, allowed values, number and date ranges, custom rules' },
    ],
  },
  {
    title: 'Reshape',
    items: [
      { key: 'columns', label: 'Column tools', hint: 'Rename, reorder, remove, split, combine or calculate columns' },
      { key: 'filter', label: 'Filter rows', hint: 'Keep rows by text, number, date or list; save the rest separately' },
      { key: 'dedupe', label: 'Remove duplicates', hint: 'Within each file or across the batch, by chosen key columns' },
    ],
  },
  {
    title: 'Outputs',
    items: [
      { key: 'split', label: 'Split into parts', hint: 'By row count, by values (e.g. one file per region) or by column groups' },
      { key: 'pivot', label: 'Summary table', hint: 'Group and total: sum, count, average, min, max, median' },
      { key: 'format', label: 'Excel extras', hint: 'Conditional formatting, charts, summary sheet, sheet protection' },
    ],
  },
];

const COMBINE = [
  { id: 'none', label: 'Keep files separate', hint: 'Each file is processed and saved on its own' },
  { id: 'merge', label: 'Merge into one', hint: 'Stack all files/sheets; columns are matched by header name' },
  { id: 'join', label: 'Join on a key', hint: 'Bring columns together by a shared ID (like VLOOKUP)' },
  { id: 'compare', label: 'Compare files', hint: 'Highlight changed cells in a copy of the first file + change log' },
] as const;

const FORMATS: { key: keyof WorkflowConfig['output']['formats']; label: string }[] = [
  { key: 'xlsx', label: 'Excel (.xlsx)' }, { key: 'csv', label: 'CSV' }, { key: 'tsv', label: 'TSV' }, { key: 'json', label: 'JSON' },
  { key: 'xml', label: 'XML' }, { key: 'html', label: 'HTML' }, { key: 'pdf', label: 'PDF' },
];

const PRESET_KEY = 'sheet-desk.presets.v1';

function loadPresets(): Record<string, WorkflowConfig> {
  try { return JSON.parse(localStorage.getItem(PRESET_KEY) || '{}'); } catch { return {}; }
}

/** Presets never contain the file list or any password. */
function cleanForPreset(c: WorkflowConfig): WorkflowConfig {
  const d = structuredClone(c);
  d.read.files = {};
  d.format.protection.password = '';
  return d;
}

export function TasksStep({ config, update, setConfig, fileCount }: Props) {
  const [presets, setPresets] = useState(loadPresets);
  const [presetName, setPresetName] = useState('');
  const [msg, setMsg] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const on = (k: ToggleKey) => config[k].enabled;

  const savePreset = () => {
    const name = presetName.trim();
    if (!name) return;
    const next = { ...presets, [name]: cleanForPreset(config) };
    try { localStorage.setItem(PRESET_KEY, JSON.stringify(next)); setPresets(next); setMsg(`Saved “${name}” in this browser.`); }
    catch { setMsg('This browser would not save the preset (storage may be blocked). Use “Download workflow” instead.'); }
    setPresetName('');
  };
  const download = () => {
    const blob = new Blob([JSON.stringify(cleanForPreset(config), null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'sheet-desk-workflow.json';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  };
  const load = async (file: File | undefined) => {
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text());
      const cfg = upgradeWorkflow(parsed);
      cfg.read.files = config.read.files;
      setConfig(cfg);
      setMsg(`Loaded workflow from ${file.name}.`);
    } catch {
      setMsg('That file is not a valid workflow (.json).');
    }
  };

  return (
    <>
      <PageHead title="What should happen to the data?" sub="Tick the tasks you need. You will set the details in the next step. Steps run in the order shown." />
      {GROUPS.map((g) => (
        <div className="card" key={g.title}>
          <h2>{g.title}</h2>
          <div className="task-grid" style={{ marginTop: 10 }}>
            {g.items.map((t) => (
              <label key={t.key} className={`task ${on(t.key) ? 'on' : ''}`}>
                <input type="checkbox" checked={on(t.key)} onChange={(e) => update((d) => { d[t.key].enabled = e.target.checked; })} />
                <span><b>{t.label}</b><small>{t.hint}</small></span>
              </label>
            ))}
          </div>
        </div>
      ))}
      <div className="card">
        <h2>Combine files?</h2>
        <p className="sub">{fileCount} file(s) ready. Combining works across all selected sheets.</p>
        <div className="task-grid" style={{ gridTemplateColumns: 'repeat(4, 1fr)' }}>
          {COMBINE.map((c) => (
            <label key={c.id} className={`task ${config.combine.mode === c.id ? 'on' : ''}`}>
              <input type="radio" name="combine" checked={config.combine.mode === c.id} onChange={() => update((d) => { d.combine.mode = c.id; })} />
              <span><b>{c.label}</b><small>{c.hint}</small></span>
            </label>
          ))}
        </div>
        {config.combine.mode !== 'none' && fileCount < 2 && config.combine.mode !== 'merge' && <Note tone="warn">This needs at least two files.</Note>}
      </div>
      <div className="card">
        <h2>Output formats</h2>
        <p className="sub">Each output is saved in every format you pick.</p>
        <div className="chips">
          {FORMATS.map((f) => (
            <button key={f.key} type="button" className={`chip ${config.output.formats[f.key] ? 'on' : ''}`} aria-pressed={config.output.formats[f.key]} onClick={() => update((d) => { d.output.formats[f.key] = !d.output.formats[f.key]; })}>{f.label}</button>
          ))}
        </div>
        <label className="check" style={{ marginTop: 12 }}>
          <input type="checkbox" checked={config.formulas.recalculate} onChange={(e) => update((d) => { d.formulas.recalculate = e.target.checked; })} />
          <span>Recalculate formulas in this app<small>Off: formulas are copied as they are and Excel recalculates them when the file is opened. On: results are recalculated here where the engine supports the functions; anything it cannot calculate is listed.</small></span>
        </label>
      </div>
      <div className="card">
        <h2>Saved workflows</h2>
        <p className="sub">Save these settings for a task you repeat. Saved workflows never include your files or passwords.</p>
        <div className="row">
          <input type="text" placeholder="Name, e.g. Monthly billing clean-up" value={presetName} onChange={(e) => setPresetName(e.target.value)} style={{ maxWidth: 320 }} />
          <button className="btn" disabled={!presetName.trim()} onClick={savePreset}>Save in this browser</button>
          <button className="btn" onClick={download}>Download workflow</button>
          <button className="btn" onClick={() => fileRef.current?.click()}>Load from file</button>
          <input ref={fileRef} type="file" accept=".json,application/json" hidden onChange={(e) => { void load(e.target.files?.[0]); e.target.value = ''; }} />
        </div>
        {Object.keys(presets).length > 0 && (
          <div className="list" style={{ marginTop: 12 }}>
            {Object.keys(presets).map((name) => (
              <div className="list-item" key={name}>
                <span className="name">{name}</span>
                <button className="btn small" onClick={() => { const c = upgradeWorkflow(presets[name]); c.read.files = config.read.files; setConfig(c); setMsg(`Loaded “${name}”.`); }}>Use</button>
                <button className="icon-btn" aria-label={`Delete ${name}`} onClick={() => { const n = { ...presets }; delete n[name]; try { localStorage.setItem(PRESET_KEY, JSON.stringify(n)); } catch { /* ignore */ } setPresets(n); }}>×</button>
              </div>
            ))}
          </div>
        )}
        {msg && <Note>{msg}</Note>}
      </div>
    </>
  );
}
