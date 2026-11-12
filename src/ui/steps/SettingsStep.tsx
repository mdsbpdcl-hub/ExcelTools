import { useMemo, useState, type ReactNode } from 'react';
import type { FileInfo } from '../../core/types';
import type {
  AggFn, ChartSpec, ConditionalRule, FilterCondition, FilterOp, FindReplaceRule, IssueAction, RuleType, ValidationRule, WorkflowConfig,
} from '../../core/workflow';
import { headerGroups } from '../../core/ops/merge';
import { RULE_LABELS } from '../../core/ops/validate';
import { AGG_LABELS } from '../../core/ops/pivot';
import { FUNCTIONS } from '../../core/expr';
import { Check, ColumnPicker, ColumnSelect, Field, Modal, More, Note, PageHead } from '../components/common';

type Update = (fn: (d: WorkflowConfig) => void) => void;
interface P { config: WorkflowConfig; update: Update; columns: string[] }

function Panel({ title, sub, children }: { title: string; sub?: string; children: ReactNode }) {
  return <div className="card"><h2>{title}</h2>{sub && <p className="sub">{sub}</p>}{children}</div>;
}

function ListEditor<T>({ items, onChange, make, render, addLabel, cols = '1fr auto' }: { items: T[]; onChange: (v: T[]) => void; make: () => T; render: (item: T, set: (patch: Partial<T>) => void, i: number) => ReactNode; addLabel: string; cols?: string }) {
  return (
    <div>
      {items.map((it, i) => (
        <div className="rule-row" key={i} style={{ gridTemplateColumns: cols }}>
          {render(it, (patch) => onChange(items.map((x, k) => (k === i ? { ...x, ...patch } : x))), i)}
          <button className="icon-btn" aria-label="Remove" onClick={() => onChange(items.filter((_, k) => k !== i))}>×</button>
        </div>
      ))}
      <button className="btn small" onClick={() => onChange([...items, make()])}>+ {addLabel}</button>
    </div>
  );
}

const DATE_ORDERS = [{ id: 'DMY', label: 'Day first (31/12/2025)' }, { id: 'MDY', label: 'Month first (12/31/2025)' }, { id: 'YMD', label: 'Year first (2025-12-31)' }] as const;
const DATE_FORMATS = ['DD-MM-YYYY', 'DD/MM/YYYY', 'YYYY-MM-DD', 'DD-MMM-YYYY', 'MM/DD/YYYY', 'DD.MM.YYYY'];

function DateOrderSelect({ value, onChange }: { value: string; onChange: (v: 'DMY' | 'MDY' | 'YMD') => void }) {
  return <select value={value} onChange={(e) => onChange(e.target.value as 'DMY')}>{DATE_ORDERS.map((o) => <option key={o.id} value={o.id}>{o.label}</option>)}</select>;
}

function ActionRow({ label, value, onChange }: { label: string; value: IssueAction; onChange: (v: IssueAction) => void }) {
  return (
    <div className="rule-row" style={{ gridTemplateColumns: '1.3fr 1fr 1fr' }}>
      <span>{label}</span>
      <select value={value.action} onChange={(e) => onChange({ ...value, action: e.target.value as IssueAction['action'] })}>
        <option value="leave">Leave unchanged</option>
        <option value="flag">Flag (highlight + list)</option>
        <option value="replace">Replace with…</option>
      </select>
      {value.action === 'replace' ? <input type="text" placeholder="(blank)" value={value.replaceWith} onChange={(e) => onChange({ ...value, replaceWith: e.target.value })} /> : <span />}
    </div>
  );
}

function QualityPanel({ config, update, columns }: P) {
  const o = config.errors;
  return (
    <Panel title="Find data problems" sub="Choose what to do with each kind of problem. Flagged cells are highlighted light red in Excel output and listed in review/Issues report.">
      <Field label="Columns to check (leave empty for all)"><ColumnPicker value={o.columns} onChange={(v) => update((d) => { d.errors.columns = v; })} columns={columns} /></Field>
      <ActionRow label="Spreadsheet errors (#N/A, #VALUE!, #REF!, #DIV/0! …)" value={o.spreadsheetErrors} onChange={(v) => update((d) => { d.errors.spreadsheetErrors = v; })} />
      <ActionRow label="Invalid dates (e.g. 31/02/2025)" value={o.invalidDates} onChange={(v) => update((d) => { d.errors.invalidDates = v; })} />
      <ActionRow label="Unexpected values (text in a number or date column)" value={o.unexpected} onChange={(v) => update((d) => { d.errors.unexpected = v; })} />
      <ActionRow label="Blank cells" value={o.blanks} onChange={(v) => update((d) => { d.errors.blanks = v; })} />
      <More><Field label="How are dates written in the data?"><DateOrderSelect value={o.dateOrder} onChange={(v) => update((d) => { d.errors.dateOrder = v; })} /></Field></More>
    </Panel>
  );
}

function CleanPanel({ config, update, columns }: P) {
  const o = config.clean;
  return (
    <Panel title="Clean up text, dates and numbers" sub="Formula cells are never changed by cleaning, so their formulas are kept.">
      <Field label="Columns to clean (leave empty for all)"><ColumnPicker value={o.columns} onChange={(v) => update((d) => { d.clean.columns = v; })} columns={columns} /></Field>
      <div className="grid2">
        <div>
          <Check checked={o.trim} onChange={(v) => update((d) => { d.clean.trim = v; })} label="Trim spaces at the start and end" />
          <Check checked={o.collapseSpaces} onChange={(v) => update((d) => { d.clean.collapseSpaces = v; })} label="Replace repeated spaces with one" />
          <Check checked={o.invisible} onChange={(v) => update((d) => { d.clean.invisible = v; })} label="Remove invisible characters" hint="Zero-width spaces, non-breaking spaces, control characters" />
        </div>
        <div>
          <Field label="Line breaks inside cells">
            <select value={o.lineBreaks} onChange={(e) => update((d) => { d.clean.lineBreaks = e.target.value as 'keep'; })}>
              <option value="keep">Keep</option><option value="space">Replace with a space</option><option value="remove">Remove</option>
            </select>
          </Field>
          <Field label="Letter case">
            <select value={o.caseMode} onChange={(e) => update((d) => { d.clean.caseMode = e.target.value as 'none'; })}>
              <option value="none">Leave as is</option><option value="upper">UPPER CASE</option><option value="lower">lower case</option><option value="title">Title Case</option><option value="sentence">Sentence case</option>
            </select>
          </Field>
        </div>
      </div>
      <Check checked={o.dates.enabled} onChange={(v) => update((d) => { d.clean.dates.enabled = v; })} label="Standardise dates" hint="Reads dates written in different ways and writes them one way" />
      {o.dates.enabled && (
        <div className="grid2" style={{ marginLeft: 24 }}>
          <Field label="Date columns"><ColumnPicker value={o.dates.columns} onChange={(v) => update((d) => { d.clean.dates.columns = v; })} columns={columns} /></Field>
          <Field label="How are dates written now?"><DateOrderSelect value={o.dates.inputOrder} onChange={(v) => update((d) => { d.clean.dates.inputOrder = v; })} /></Field>
          <Field label="Write dates as">
            <select value={o.dates.asText ? o.dates.outputFormat : 'real'} onChange={(e) => update((d) => { if (e.target.value === 'real') d.clean.dates.asText = false; else { d.clean.dates.asText = true; d.clean.dates.outputFormat = e.target.value; } })}>
              <option value="real">Real Excel dates (recommended)</option>
              {DATE_FORMATS.map((f) => <option key={f} value={f}>Text: {f}</option>)}
            </select>
          </Field>
        </div>
      )}
      <Check checked={o.numbers.enabled} onChange={(v) => update((d) => { d.clean.numbers.enabled = v; })} label="Turn numbers stored as text into numbers" hint="Handles 1,23,456 · ₹ 2,500 · (1,000) · 12%" />
      {o.numbers.enabled && <div style={{ marginLeft: 24 }}><Field label="Number columns"><ColumnPicker value={o.numbers.columns} onChange={(v) => update((d) => { d.clean.numbers.columns = v; })} columns={columns} /></Field></div>}
    </Panel>
  );
}

const newFR = (): FindReplaceRule => ({ find: '', replace: '', mode: 'partial', caseSensitive: false, columns: [], sheets: [], files: [] });

function FindReplacePanel({ config, update, columns }: P) {
  return (
    <Panel title="Find and replace" sub="Rules run in order. Formula cells are skipped so formulas are kept. In pattern mode, use $1, $2 to reuse matched parts.">
      <ListEditor<FindReplaceRule>
        items={config.findReplace.rules}
        onChange={(v) => update((d) => { d.findReplace.rules = v; })}
        make={newFR}
        addLabel="Add rule"
        cols="1fr auto"
        render={(r, set) => (
          <div>
            <div className="grid3">
              <input type="text" placeholder="Find" value={r.find} onChange={(e) => set({ find: e.target.value })} />
              <input type="text" placeholder="Replace with (empty = remove)" value={r.replace} onChange={(e) => set({ replace: e.target.value })} />
              <select value={r.mode} onChange={(e) => set({ mode: e.target.value as FindReplaceRule['mode'] })}>
                <option value="partial">Anywhere in the cell</option><option value="exact">Whole cell only</option><option value="regex">Pattern (regular expression)</option>
              </select>
            </div>
            <Check checked={r.caseSensitive} onChange={(v) => set({ caseSensitive: v })} label="Match upper/lower case exactly" />
            <More label="Limit to columns, sheets or files">
              <div className="grid3">
                <Field label="Columns (empty = all)"><ColumnPicker value={r.columns} onChange={(v) => set({ columns: v })} columns={columns} /></Field>
                <Field label="Sheet names (empty = all)"><ColumnPicker value={r.sheets} onChange={(v) => set({ sheets: v })} columns={[]} placeholder="Type a sheet name" /></Field>
                <Field label="File names (empty = all)"><ColumnPicker value={r.files} onChange={(v) => set({ files: v })} columns={[]} placeholder="e.g. bills.xlsx" /></Field>
              </div>
            </More>
          </div>
        )}
      />
    </Panel>
  );
}

function ExpressionHelp({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="Writing a calculation" onClose={onClose}>
      <p>Refer to columns as <span className="mono">[Column Name]</span>. Use <span className="mono">+ - * / ^</span> for maths, <span className="mono">&amp;</span> to join text and <span className="mono">= &lt;&gt; &lt; &gt; &lt;= &gt;=</span> to compare. Text goes in quotes.</p>
      <p className="mono">[Units] * [Rate]<br />IF([Amount] &gt;= 10000, "High", "Normal")<br />PROPER([Name]) &amp; " (" &amp; [Consumer No] &amp; ")"<br />DAYS(TODAY(), [Bill Date])</p>
      <table className="data"><thead><tr><th>Function</th><th>How to use</th></tr></thead><tbody>
        {Object.entries(FUNCTIONS).map(([k, f]) => <tr key={k}><td className="mono">{k}</td><td>{f.help}</td></tr>)}
      </tbody></table>
      <p className="muted small">These calculations are worked out by this app and saved as values. They are not Excel formulas.</p>
    </Modal>
  );
}

function ColumnsPanel({ config, update, columns }: P) {
  const o = config.columns;
  const [help, setHelp] = useState(false);
  return (
    <Panel title="Column tools" sub="Applied in this order: split → combine → new columns → rename → remove → order. Changing columns means formulas in those sheets are saved as values.">
      <Field label="Rename columns">
        <ListEditor items={o.rename} onChange={(v) => update((d) => { d.columns.rename = v; })} make={() => ({ from: '', to: '' })} addLabel="Rename a column" cols="1fr 1fr auto"
          render={(r, set) => <><ColumnSelect value={r.from} onChange={(v) => set({ from: v })} columns={columns} placeholder="Current name" /><input type="text" placeholder="New name" value={r.to} onChange={(e) => set({ to: e.target.value })} /></>} />
      </Field>
      <Field label="Remove columns"><ColumnPicker value={o.remove} onChange={(v) => update((d) => { d.columns.remove = v; })} columns={columns} /></Field>
      <Field label="Put these columns first, in this order"><ColumnPicker value={o.order} onChange={(v) => update((d) => { d.columns.order = v; })} columns={columns} /></Field>
      <Check checked={o.keepOnlyListed} onChange={(v) => update((d) => { d.columns.keepOnlyListed = v; })} label="Keep only the columns listed above (remove all others)" />
      <More label="Split, combine or calculate columns" open={!!(o.split.length || o.combine.length || o.derived.length)}>
        <Field label="Split a column (e.g. “Ram Kumar” → First, Last)">
          <ListEditor items={o.split} onChange={(v) => update((d) => { d.columns.split = v; })} make={() => ({ column: '', delimiter: ',', into: [], keepOriginal: false })} addLabel="Split a column" cols="1fr 110px 1.4fr auto auto"
            render={(r, set) => <>
              <ColumnSelect value={r.column} onChange={(v) => set({ column: v })} columns={columns} />
              <input type="text" placeholder="Split at" value={r.delimiter} onChange={(e) => set({ delimiter: e.target.value })} title="Use \t for tab" />
              <input type="text" placeholder="New names: First, Last" value={r.into.join(', ')} onChange={(e) => set({ into: e.target.value.split(',').map((s) => s.trim()) })} />
              <Check checked={r.keepOriginal} onChange={(v) => set({ keepOriginal: v })} label="Keep original" />
            </>} />
        </Field>
        <Field label="Combine columns (e.g. House, Street, City → Address)">
          <ListEditor items={o.combine} onChange={(v) => update((d) => { d.columns.combine = v; })} make={() => ({ columns: [], separator: ' ', into: '', keepOriginals: true, skipBlanks: true })} addLabel="Combine columns" cols="1.6fr 90px 1fr auto auto"
            render={(r, set) => <>
              <ColumnPicker value={r.columns} onChange={(v) => set({ columns: v })} columns={columns} />
              <input type="text" placeholder="Between" value={r.separator} onChange={(e) => set({ separator: e.target.value })} />
              <input type="text" placeholder="New column name" value={r.into} onChange={(e) => set({ into: e.target.value })} />
              <Check checked={r.keepOriginals} onChange={(v) => set({ keepOriginals: v })} label="Keep originals" />
            </>} />
        </Field>
        <Field label="New calculated column">
          <ListEditor items={o.derived} onChange={(v) => update((d) => { d.columns.derived = v; })} make={() => ({ name: '', expression: '' })} addLabel="Add calculated column" cols="1fr 2fr auto"
            render={(r, set) => <>
              <input type="text" placeholder="Column name" value={r.name} onChange={(e) => set({ name: e.target.value })} />
              <input type="text" className="mono" placeholder='e.g. [Units] * [Rate]' value={r.expression} onChange={(e) => set({ expression: e.target.value })} />
            </>} />
        </Field>
        <button className="btn link" onClick={() => setHelp(true)}>How to write a calculation</button>
      </More>
      {help && <ExpressionHelp onClose={() => setHelp(false)} />}
    </Panel>
  );
}

const OPS: { id: FilterOp; label: string; values: 0 | 1 | 2 }[] = [
  { id: 'equals', label: 'is', values: 1 }, { id: 'notEquals', label: 'is not', values: 1 },
  { id: 'contains', label: 'contains', values: 1 }, { id: 'notContains', label: 'does not contain', values: 1 },
  { id: 'startsWith', label: 'starts with', values: 1 }, { id: 'endsWith', label: 'ends with', values: 1 },
  { id: 'inList', label: 'is one of (comma list)', values: 1 }, { id: 'notInList', label: 'is not one of', values: 1 },
  { id: 'gt', label: '> greater than', values: 1 }, { id: 'gte', label: '≥ at least', values: 1 }, { id: 'lt', label: '< less than', values: 1 }, { id: 'lte', label: '≤ at most', values: 1 },
  { id: 'between', label: 'number between', values: 2 },
  { id: 'dateOn', label: 'date is', values: 1 }, { id: 'dateBefore', label: 'date before', values: 1 }, { id: 'dateAfter', label: 'date after', values: 1 }, { id: 'dateBetween', label: 'date between', values: 2 },
  { id: 'blank', label: 'is blank', values: 0 }, { id: 'notBlank', label: 'is not blank', values: 0 },
  { id: 'regex', label: 'matches pattern', values: 1 }, { id: 'expression', label: 'custom rule (expression)', values: 1 },
];

function FilterPanel({ config, update, columns }: P) {
  const o = config.filter;
  return (
    <Panel title="Filter rows" sub="Rows that match are kept. Dates can be typed as DD-MM-YYYY.">
      <div className="row" style={{ marginBottom: 8 }}>
        <span>Keep rows that match</span>
        <select value={o.match} onChange={(e) => update((d) => { d.filter.match = e.target.value as 'all'; })} style={{ width: 'auto' }}>
          <option value="all">all conditions</option><option value="any">any condition</option>
        </select>
      </div>
      <ListEditor<FilterCondition> items={o.conditions} onChange={(v) => update((d) => { d.filter.conditions = v; })} addLabel="Add condition" cols="1fr 1fr 1fr 1fr auto"
        make={() => ({ column: '', op: 'equals', value: '', value2: '', caseSensitive: false })}
        render={(c, set) => {
          const op = OPS.find((x) => x.id === c.op)!;
          return <>
            {c.op === 'expression' ? <span className="muted small">Expression:</span> : <ColumnSelect value={c.column} onChange={(v) => set({ column: v })} columns={columns} />}
            <select value={c.op} onChange={(e) => set({ op: e.target.value as FilterOp })}>{OPS.map((x) => <option key={x.id} value={x.id}>{x.label}</option>)}</select>
            {op.values >= 1 ? <input type="text" className={c.op === 'expression' ? 'mono' : ''} placeholder={c.op === 'expression' ? 'AND([Units] > 100, [Circle] = "Patna")' : op.values === 2 ? 'From' : 'Value'} value={c.value} onChange={(e) => set({ value: e.target.value })} /> : <span />}
            {op.values === 2 ? <input type="text" placeholder="To" value={c.value2} onChange={(e) => set({ value2: e.target.value })} /> : <span />}
          </>;
        }} />
      <Check checked={o.saveNonMatching} onChange={(v) => update((d) => { d.filter.saveNonMatching = v; })} label="Save non-matching rows as a separate file" hint="Otherwise rows that do not match are left out of the output." />
      <More><Field label="How are dates written in the data?"><DateOrderSelect value={o.dateOrder} onChange={(v) => update((d) => { d.filter.dateOrder = v; })} /></Field></More>
    </Panel>
  );
}

function DedupePanel({ config, update, columns }: P) {
  const o = config.dedupe;
  return (
    <Panel title="Remove duplicates" sub="Rows are duplicates when the key columns match.">
      <Field label="Key columns (leave empty to compare whole rows)"><ColumnPicker value={o.keys} onChange={(v) => update((d) => { d.dedupe.keys = v; })} columns={columns} /></Field>
      <div className="grid2">
        <Field label="Which copy to keep">
          <select value={o.keep} onChange={(e) => update((d) => { d.dedupe.keep = e.target.value as 'first'; })}>
            <option value="first">The first one</option><option value="last">The last one</option><option value="mostComplete">The most complete one (fewest blanks)</option>
          </select>
        </Field>
        <Field label="Look for duplicates">
          <select value={o.scope} onChange={(e) => update((d) => { d.dedupe.scope = e.target.value as 'file'; })}>
            <option value="file">Within each file/sheet</option><option value="batch">Across all files (in the order added)</option>
          </select>
        </Field>
      </div>
      <Check checked={o.exportRemoved} onChange={(v) => update((d) => { d.dedupe.exportRemoved = v; })} label="Save removed duplicates to review/ with their source file and row" />
      <More>
        <Check checked={o.ignoreCase} onChange={(v) => update((d) => { d.dedupe.ignoreCase = v; })} label="Ignore upper/lower case when comparing" />
        <Check checked={o.trimSpaces} onChange={(v) => update((d) => { d.dedupe.trimSpaces = v; })} label="Ignore extra spaces when comparing" />
      </More>
    </Panel>
  );
}

const newRule = (): ValidationRule => ({ column: '', type: 'required', values: '', min: '', max: '', pattern: '', expression: '', message: '' });

function ValidatePanel({ config, update, columns }: P) {
  const o = config.validate;
  return (
    <Panel title="Check rules" sub="Outputs with rows that break a rule go to review/, with the problem cells highlighted and listed in the issues report.">
      <ListEditor<ValidationRule> items={o.rules} onChange={(v) => update((d) => { d.validate.rules = v; })} make={newRule} addLabel="Add rule" cols="1fr 1fr 1.6fr auto"
        render={(r, set) => <>
          {r.type === 'expression' ? <span className="muted small">Whole row</span> : <ColumnSelect value={r.column} onChange={(v) => set({ column: v })} columns={columns} />}
          <select value={r.type} onChange={(e) => set({ type: e.target.value as RuleType })}>{Object.entries(RULE_LABELS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
          <div className="row">
            {r.type === 'list' && <input type="text" placeholder="Allowed values, comma separated" value={r.values} onChange={(e) => set({ values: e.target.value })} />}
            {(r.type === 'range' || r.type === 'dateRange') && <>
              <input type="text" placeholder={r.type === 'range' ? 'Min' : 'From DD-MM-YYYY'} value={r.min} onChange={(e) => set({ min: e.target.value })} style={{ flex: 1 }} />
              <input type="text" placeholder={r.type === 'range' ? 'Max' : 'To DD-MM-YYYY'} value={r.max} onChange={(e) => set({ max: e.target.value })} style={{ flex: 1 }} />
            </>}
            {r.type === 'pattern' && <input type="text" className="mono" placeholder="e.g. ^[0-9]{10}$ for 10 digits" value={r.pattern} onChange={(e) => set({ pattern: e.target.value })} />}
            {r.type === 'expression' && <input type="text" className="mono" placeholder='e.g. [Paid] <= [Billed]' value={r.expression} onChange={(e) => set({ expression: e.target.value })} />}
            <input type="text" placeholder="Message (optional)" value={r.message} onChange={(e) => set({ message: e.target.value })} />
          </div>
        </>} />
      <Check checked={o.markInvalidCells} onChange={(v) => update((d) => { d.validate.markInvalidCells = v; })} label="Highlight invalid cells in Excel output" />
      <Check checked={o.addDropdowns} onChange={(v) => update((d) => { d.validate.addDropdowns = v; })} label="Add dropdown lists to Excel output for “one of a list” rules" />
      <More><Field label="How are dates written in the data?"><DateOrderSelect value={o.dateOrder} onChange={(v) => update((d) => { d.validate.dateOrder = v; })} /></Field></More>
    </Panel>
  );
}

function CombinePanel({ config, update, columns, infos }: P & { infos: FileInfo[] }) {
  const c = config.combine;
  const groups = useMemo(() => headerGroups(infos.filter((i) => i.status === 'ready').flatMap((i) => i.sheets.filter((s) => (config.read.files[i.id]?.sheets ? config.read.files[i.id].sheets!.includes(s.name) : !s.hidden || config.read.includeHidden)))), [infos, config.read]);
  if (c.mode === 'merge') {
    const fileCount = infos.filter((i) => i.status === 'ready').length;
    return (
      <Panel title="Merge into one" sub="Columns with the same name (ignoring case, spaces and punctuation) are put together. To join two different names, type the same output name for both.">
        <div className="table-wrap" style={{ maxHeight: 300 }}>
          <table className="data"><thead><tr><th>Found as</th><th>In files</th><th>Output column name</th></tr></thead><tbody>
            {groups.map((g) => (
              <tr key={g.key}>
                <td title={g.spellings.join(' | ')}>{g.spellings.join(' | ')}</td>
                <td>{g.files} of {fileCount}</td>
                <td><input type="text" value={c.merge.headerMap[g.key] ?? g.suggested} onChange={(e) => update((d) => { d.combine.merge.headerMap[g.key] = e.target.value; })} /></td>
              </tr>
            ))}
          </tbody></table>
        </div>
        <div className="row" style={{ marginTop: 10 }}>
          <Check checked={c.merge.addSourceFile} onChange={(v) => update((d) => { d.combine.merge.addSourceFile = v; })} label="Add “Source file” column" />
          <Check checked={c.merge.addSourceSheet} onChange={(v) => update((d) => { d.combine.merge.addSourceSheet = v; })} label="Add “Source sheet”" />
          <Check checked={c.merge.addSourceRow} onChange={(v) => update((d) => { d.combine.merge.addSourceRow = v; })} label="Add “Source row”" />
        </div>
        <div className="grid2">
          <Field label="Output name"><input type="text" value={c.merge.outputName} onChange={(e) => update((d) => { d.combine.merge.outputName = e.target.value; })} /></Field>
          <div style={{ paddingTop: 22 }}><Check checked={c.keepIndividual} onChange={(v) => update((d) => { d.combine.keepIndividual = v; })} label="Also save each processed file separately" /></div>
        </div>
      </Panel>
    );
  }
  if (c.mode === 'join') {
    return (
      <Panel title="Join on a key" sub="The first file is the starting point; columns from the other files are added where the key matches. Rows without a match and keys that repeat are listed in review/.">
        <Field label="Key column(s) — must have the same name in every file"><ColumnPicker value={c.join.keys} onChange={(v) => update((d) => { d.combine.join.keys = v; })} columns={columns} /></Field>
        <div className="grid2">
          <Field label="Which rows to keep">
            <select value={c.join.type} onChange={(e) => update((d) => { d.combine.join.type = e.target.value as 'left'; })}>
              <option value="left">All rows of the first file (like VLOOKUP)</option><option value="inner">Only rows found in every file</option><option value="full">All rows from all files</option>
            </select>
          </Field>
          <Field label="Output name"><input type="text" value={c.join.outputName} onChange={(e) => update((d) => { d.combine.join.outputName = e.target.value; })} /></Field>
        </div>
        <Check checked={c.join.ignoreCase} onChange={(v) => update((d) => { d.combine.join.ignoreCase = v; })} label="Ignore upper/lower case in keys" />
        <Check checked={c.join.trimSpaces} onChange={(v) => update((d) => { d.combine.join.trimSpaces = v; })} label="Ignore spaces around keys" />
        <Check checked={c.keepIndividual} onChange={(v) => update((d) => { d.combine.keepIndividual = v; })} label="Also save each processed file separately" />
      </Panel>
    );
  }
  if (c.mode === 'compare') {
    return (
      <Panel title="Compare files" sub="The first file added is the baseline. You get (1) a copy of it with changed cells highlighted light yellow — each with a note showing the other value — and (2) a change log with old and new values, file, sheet, row and column.">
        <Field label="Match rows by key column(s) — leave empty to compare row by row"><ColumnPicker value={c.compare.keys} onChange={(v) => update((d) => { d.combine.compare.keys = v; })} columns={columns} /></Field>
        <div className="grid2">
          <Field label="Ignore number differences up to"><input type="number" min={0} step="any" value={c.compare.numericTolerance} onChange={(e) => update((d) => { d.combine.compare.numericTolerance = Number(e.target.value) || 0; })} /></Field>
          <Field label="Output name"><input type="text" value={c.compare.outputName} onChange={(e) => update((d) => { d.combine.compare.outputName = e.target.value; })} /></Field>
        </div>
        <Check checked={c.compare.ignoreCase} onChange={(v) => update((d) => { d.combine.compare.ignoreCase = v; })} label="Ignore upper/lower case" />
        <Check checked={c.compare.trimSpaces} onChange={(v) => update((d) => { d.combine.compare.trimSpaces = v; })} label="Ignore extra spaces" />
        <Note>With several files, each is compared with the first. Sheets are paired by name (or one-to-one when each file has a single sheet).</Note>
      </Panel>
    );
  }
  return null;
}

function SplitPanel({ config, update, columns }: P) {
  const o = config.split;
  return (
    <Panel title="Split into parts">
      <div className="chips" style={{ marginBottom: 12 }}>
        {(['values', 'rows', 'columns'] as const).map((m) => (
          <button key={m} className={`chip ${o.mode === m ? 'on' : ''}`} onClick={() => update((d) => { d.split.mode = m; })}>
            {m === 'values' ? 'By values in a column' : m === 'rows' ? 'By number of rows' : 'By column groups'}
          </button>
        ))}
      </div>
      {o.mode === 'values' && <Field label="One part for each value of (e.g. Circle, Division)"><ColumnPicker value={o.columns} onChange={(v) => update((d) => { d.split.columns = v; })} columns={columns} /></Field>}
      {o.mode === 'rows' && <Field label="Rows per part"><input type="number" min={1} value={o.rowCount} onChange={(e) => update((d) => { d.split.rowCount = Math.max(1, Number(e.target.value) || 1); })} /></Field>}
      {o.mode === 'columns' && <>
        <Field label="Column groups — separate groups with ; (e.g. Name, Phone; Address, City)"><input type="text" value={o.columnGroups} onChange={(e) => update((d) => { d.split.columnGroups = e.target.value; })} /></Field>
        <Field label="Repeat these columns in every part (e.g. an ID)"><ColumnPicker value={o.repeatColumns} onChange={(v) => update((d) => { d.split.repeatColumns = v; })} columns={columns} /></Field>
      </>}
      <div className="grid2">
        <Field label="Save parts as">
          <select value={o.target} onChange={(e) => update((d) => { d.split.target = e.target.value as 'files'; })}>
            <option value="files">Separate files</option><option value="sheets">Sheets in one workbook</option>
          </select>
        </Field>
        {o.target === 'files' && <Field label="File name rule" hint="{file} original name · {value} part value · {n} part number · {sheet} · {date}"><input type="text" value={o.naming} onChange={(e) => update((d) => { d.split.naming = e.target.value; })} /></Field>}
      </div>
    </Panel>
  );
}

function PivotPanel({ config, update, columns }: P) {
  const o = config.pivot;
  return (
    <Panel title="Summary table" sub="Saved as a separate output next to the data, in the formats you chose.">
      <div className="grid2">
        <Field label="Group rows by"><ColumnPicker value={o.groupBy} onChange={(v) => update((d) => { d.pivot.groupBy = v; })} columns={columns} /></Field>
        <Field label="Spread across columns by (optional)"><ColumnSelect value={o.pivotColumn} onChange={(v) => update((d) => { d.pivot.pivotColumn = v; })} columns={columns} placeholder="e.g. Month" /></Field>
      </div>
      <Field label="Values">
        <ListEditor items={o.values} onChange={(v) => update((d) => { d.pivot.values = v; })} make={() => ({ column: '', agg: 'sum' as AggFn })} addLabel="Add value" cols="1fr 1fr auto"
          render={(v, set) => <>
            <select value={v.agg} onChange={(e) => set({ agg: e.target.value as AggFn })}>{Object.entries(AGG_LABELS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
            <ColumnSelect value={v.column} onChange={(c) => set({ column: c })} columns={columns} />
          </>} />
      </Field>
      <div className="grid2">
        <Check checked={o.includeTotals} onChange={(v) => update((d) => { d.pivot.includeTotals = v; })} label="Add totals" />
        <Field label="Name"><input type="text" value={o.outputName} onChange={(e) => update((d) => { d.pivot.outputName = e.target.value; })} /></Field>
      </div>
    </Panel>
  );
}

const CF_TYPES: { id: ConditionalRule['type']; label: string; values: number }[] = [
  { id: 'greaterThan', label: 'Highlight if greater than', values: 1 }, { id: 'lessThan', label: 'Highlight if less than', values: 1 },
  { id: 'between', label: 'Highlight if between', values: 2 }, { id: 'equal', label: 'Highlight if equal to', values: 1 },
  { id: 'containsText', label: 'Highlight if text contains', values: 1 }, { id: 'duplicates', label: 'Highlight duplicates', values: 0 },
  { id: 'dataBar', label: 'Data bars', values: 0 }, { id: 'colorScale', label: 'Colour scale (red–yellow–green)', values: 0 }, { id: 'iconSet', label: 'Traffic-light icons', values: 0 },
];

function FormatPanel({ config, update, columns }: P) {
  const o = config.format;
  return (
    <Panel title="Excel extras" sub="These apply to Excel (.xlsx) output only.">
      <Field label="Conditional formatting">
        <ListEditor<ConditionalRule> items={o.conditional} onChange={(v) => update((d) => { d.format.conditional = v; })} addLabel="Add highlight rule" cols="1fr 1.3fr 1fr 1fr 70px auto"
          make={() => ({ column: '', type: 'greaterThan', value: '', value2: '', color: 'FFD7D7' })}
          render={(r, set) => {
            const t = CF_TYPES.find((x) => x.id === r.type)!;
            return <>
              <ColumnSelect value={r.column} onChange={(v) => set({ column: v })} columns={columns} />
              <select value={r.type} onChange={(e) => set({ type: e.target.value as ConditionalRule['type'] })}>{CF_TYPES.map((x) => <option key={x.id} value={x.id}>{x.label}</option>)}</select>
              {t.values >= 1 ? <input type="text" placeholder="Value" value={r.value} onChange={(e) => set({ value: e.target.value })} /> : <span />}
              {t.values === 2 ? <input type="text" placeholder="To" value={r.value2} onChange={(e) => set({ value2: e.target.value })} /> : <span />}
              {r.type === 'colorScale' || r.type === 'iconSet' ? <span /> : <input type="color" aria-label="Colour" value={`#${r.color}`} onChange={(e) => set({ color: e.target.value.slice(1).toUpperCase() })} style={{ padding: 0, height: 34 }} />}
            </>;
          }} />
      </Field>
      <Field label="Charts">
        <ListEditor<ChartSpec> items={o.charts} onChange={(v) => update((d) => { d.format.charts = v; })} addLabel="Add chart" cols="1fr 1fr 1fr 1fr 1.2fr auto"
          make={() => ({ type: 'bar', category: '', value: '', agg: 'sum', title: '', maxCategories: 12 })}
          render={(c, set) => <>
            <select value={c.type} onChange={(e) => set({ type: e.target.value as ChartSpec['type'] })}>
              <option value="bar">Bar</option><option value="line">Line</option><option value="pie">Pie</option><option value="scatter">Scatter</option><option value="waterfall">Waterfall</option>
            </select>
            <ColumnSelect value={c.category} onChange={(v) => set({ category: v })} columns={columns} placeholder={c.type === 'scatter' ? 'X column' : 'Category'} />
            {c.type === 'scatter' ? <span className="muted small">vs</span> : <select value={c.agg} onChange={(e) => set({ agg: e.target.value as AggFn })}>{Object.entries(AGG_LABELS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>}
            <ColumnSelect value={c.value} onChange={(v) => set({ value: v })} columns={columns} placeholder={c.type === 'scatter' ? 'Y column' : 'Value'} />
            <input type="text" placeholder="Title (optional)" value={c.title} onChange={(e) => set({ title: e.target.value })} />
          </>} />
        <p className="muted small">Charts are added to a Summary sheet as pictures. The Excel library used cannot create native, editable Excel charts.</p>
      </Field>
      <Check checked={o.summarySheet} onChange={(v) => update((d) => { d.format.summarySheet = v; })} label="Add a Summary sheet with key numbers" />
      <Check checked={o.profileSheet} onChange={(v) => update((d) => { d.format.profileSheet = v; })} label="Add a Data profile sheet (types, blanks, unique values per column)" />
      <More label="Layout and dates">
        <Check checked={o.styleHeader} onChange={(v) => update((d) => { d.format.styleHeader = v; })} label="Style the header row (bold, light green)" />
        <Check checked={o.freezeHeader} onChange={(v) => update((d) => { d.format.freezeHeader = v; })} label="Freeze the header row" />
        <Check checked={o.autoFilter} onChange={(v) => update((d) => { d.format.autoFilter = v; })} label="Add filter buttons to the header" />
        <Field label="Date format in Excel"><select value={o.excelDateFormat} onChange={(e) => update((d) => { d.format.excelDateFormat = e.target.value; })}>{['dd-mm-yyyy', 'dd/mm/yyyy', 'yyyy-mm-dd', 'dd-mmm-yyyy', 'mm/dd/yyyy'].map((f) => <option key={f}>{f}</option>)}</select></Field>
        <p className="muted small">When original formatting is kept, the header style, freeze and filter options are not applied so the file looks as before.</p>
      </More>
      <More label="Sheet protection" open={o.protection.enabled}>
        <Note tone="warn">Sheet protection stops accidental edits in Excel. It is <b>not encryption</b>: anyone can still open and read the file, and protection can be removed with common tools. This app cannot add a password to open the file.</Note>
        <Check checked={o.protection.enabled} onChange={(v) => update((d) => { d.format.protection.enabled = v; })} label="Protect sheets" />
        {o.protection.enabled && <>
          <div className="grid2">
            <Field label="Password to unprotect (optional, not saved in workflows)"><input type="password" autoComplete="new-password" value={o.protection.password} onChange={(e) => update((d) => { d.format.protection.password = e.target.value; })} /></Field>
            <Field label="Columns people may still edit"><ColumnPicker value={o.protection.unlockedColumns} onChange={(v) => update((d) => { d.format.protection.unlockedColumns = v; })} columns={columns} /></Field>
          </div>
          <div className="row">
            <Check checked={o.protection.allowSort} onChange={(v) => update((d) => { d.format.protection.allowSort = v; })} label="Allow sorting" />
            <Check checked={o.protection.allowFilter} onChange={(v) => update((d) => { d.format.protection.allowFilter = v; })} label="Allow filtering" />
            <Check checked={o.protection.allowFormatColumns} onChange={(v) => update((d) => { d.format.protection.allowFormatColumns = v; })} label="Allow column widths" />
          </div>
        </>}
      </More>
    </Panel>
  );
}

function OutputPanel({ config, update, columns }: P) {
  const o = config.output;
  const f = o.formats;
  return (
    <Panel title="Output and package" sub="Results are packed into a ZIP with success/, review/ and error/ folders, plus run-summary.json and run-report.txt.">
      <div className="grid2">
        <Field label="Package name"><input type="text" value={o.packageName} onChange={(e) => update((d) => { d.output.packageName = e.target.value; })} /></Field>
        <Field label="Add to every output name (optional)" hint="e.g. _clean → bills_clean.xlsx"><input type="text" value={o.nameSuffix} onChange={(e) => update((d) => { d.output.nameSuffix = e.target.value; })} /></Field>
      </div>
      <Check checked={o.keepFolders} onChange={(v) => update((d) => { d.output.keepFolders = v; })} label="Keep the original sub-folder structure inside each folder" />
      <Check checked={o.includeFailedOriginals} onChange={(v) => update((d) => { d.output.includeFailedOriginals = v; })} label="Copy files that failed into error/ (unchanged)" />
      {f.xlsx && <Check checked={o.keepOriginalFormatting} onChange={(v) => update((d) => { d.output.keepOriginalFormatting = v; })} label="Keep original Excel formatting when possible" hint="For .xlsx files whose rows and columns were not moved. Charts, pivot tables and macros in the original are not kept by the Excel library used." />}
      {(f.csv || f.tsv || f.json || f.xml || f.html || f.pdf) && (
        <More label="Text, PDF and date settings" open={f.pdf}>
          <div className="grid3">
            {(f.csv || f.tsv) && <Field label="CSV/TSV encoding">
              <select value={o.csvEncoding} onChange={(e) => update((d) => { d.output.csvEncoding = e.target.value as 'utf-8'; })}>
                <option value="utf-8-bom">UTF-8 with BOM (opens correctly in Excel)</option><option value="utf-8">UTF-8</option><option value="utf-16le">UTF-16 (Unicode text)</option><option value="windows-1252">Windows-1252 (Western, no Hindi)</option>
              </select>
            </Field>}
            {f.csv && <Field label="CSV delimiter"><select value={o.csvDelimiter} onChange={(e) => update((d) => { d.output.csvDelimiter = e.target.value; })}><option value=",">Comma</option><option value=";">Semicolon</option><option value="|">Pipe</option></select></Field>}
            <Field label="Dates in text outputs"><select value={o.textDateFormat} onChange={(e) => update((d) => { d.output.textDateFormat = e.target.value; })}>{DATE_FORMATS.map((x) => <option key={x}>{x}</option>)}</select></Field>
            {f.json && <Field label="JSON layout"><select value={o.jsonStyle} onChange={(e) => update((d) => { d.output.jsonStyle = e.target.value as 'records'; })}><option value="records">List of rows with column names</option><option value="arrays">Columns + rows (compact)</option></select></Field>}
          </div>
          {f.pdf && <>
            <div className="field-label" style={{ marginTop: 8 }}>PDF</div>
            <div className="grid3">
              <Field label="Orientation"><select value={o.pdf.orientation} onChange={(e) => update((d) => { d.output.pdf.orientation = e.target.value as 'portrait'; })}><option value="landscape">Landscape</option><option value="portrait">Portrait</option></select></Field>
              <Field label="Paper size"><select value={o.pdf.paper} onChange={(e) => update((d) => { d.output.pdf.paper = e.target.value as 'a4'; })}><option value="a4">A4</option><option value="a3">A3</option><option value="letter">Letter</option><option value="legal">Legal</option></select></Field>
              <Field label="Scaling"><select value={o.pdf.scale} onChange={(e) => update((d) => { d.output.pdf.scale = e.target.value as 'fitWidth'; })}><option value="fitWidth">Fit columns to page width</option><option value="actual">Actual size (wrap wide tables)</option></select></Field>
              <Field label="Font size (pt)"><input type="number" min={4} max={14} value={o.pdf.fontSize} onChange={(e) => update((d) => { d.output.pdf.fontSize = Math.min(14, Math.max(4, Number(e.target.value) || 8)); })} /></Field>
              <Field label="Margins (mm)"><input type="number" min={3} max={40} value={o.pdf.marginMm} onChange={(e) => update((d) => { d.output.pdf.marginMm = Math.min(40, Math.max(3, Number(e.target.value) || 10)); })} /></Field>
              <Field label="Rows per page (0 = automatic)"><input type="number" min={0} value={o.pdf.rowsPerPage} onChange={(e) => update((d) => { d.output.pdf.rowsPerPage = Math.max(0, Number(e.target.value) || 0); })} /></Field>
              <Field label="New page when this column changes"><ColumnSelect value={o.pdf.breakOnColumn} onChange={(v) => update((d) => { d.output.pdf.breakOnColumn = v; })} columns={columns} placeholder="(none)" /></Field>
              <Field label="Title on each page"><input type="text" value={o.pdf.title} onChange={(e) => update((d) => { d.output.pdf.title = e.target.value; })} /></Field>
            </div>
            <Check checked={o.pdf.repeatHeader} onChange={(v) => update((d) => { d.output.pdf.repeatHeader = v; })} label="Repeat the header row on every page" />
            <Note tone="warn">PDF fonts here cover English/Latin text only. For Hindi or other scripts, also choose HTML and use the browser’s Print › Save as PDF (page size and orientation are already set).</Note>
          </>}
        </More>
      )}
    </Panel>
  );
}

export function SettingsStep({ config, update, columns, infos }: P & { infos: FileInfo[] }) {
  const any = ['errors', 'clean', 'findReplace', 'columns', 'filter', 'dedupe', 'validate', 'split', 'pivot', 'format'].some((k) => (config as unknown as Record<string, { enabled: boolean }>)[k].enabled) || config.combine.mode !== 'none';
  const p = { config, update, columns };
  return (
    <>
      <PageHead title="Set the details" sub={columns.length ? `Column suggestions come from the ${columns.length} column names found in your files. You can also type a name.` : 'Add files first to get column suggestions.'} />
      {!any && <Note>No tasks are ticked, so files will only be converted to the chosen output formats. Go back to add tasks.</Note>}
      {config.errors.enabled && <QualityPanel {...p} />}
      {config.clean.enabled && <CleanPanel {...p} />}
      {config.findReplace.enabled && <FindReplacePanel {...p} />}
      {config.columns.enabled && <ColumnsPanel {...p} />}
      {config.filter.enabled && <FilterPanel {...p} />}
      {config.dedupe.enabled && <DedupePanel {...p} />}
      {config.combine.mode !== 'none' && <CombinePanel {...p} infos={infos} />}
      {config.split.enabled && <SplitPanel {...p} />}
      {config.pivot.enabled && <PivotPanel {...p} />}
      {config.validate.enabled && <ValidatePanel {...p} />}
      {config.format.enabled && <FormatPanel {...p} />}
      <OutputPanel {...p} />
    </>
  );
}
