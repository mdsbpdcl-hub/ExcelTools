import { useState } from 'react';
import type { FileInfo } from '../../core/types';
import type { FileReadOverride } from '../../core/workflow';
import { DELIMITERS, ENCODINGS } from '../../core/parse/csv';
import { DataPreview, Field, Modal, Note, Tag } from '../components/common';

interface Props {
  info: FileInfo;
  override: FileReadOverride;
  includeHidden: boolean;
  onClose: () => void;
  onSave: (ov: FileReadOverride) => void;
}

export function FileSettings({ info, override, includeHidden, onClose, onSave }: Props) {
  const isText = info.kind === 'csv' || info.kind === 'tsv' || info.kind === 'txt';
  const defaultSheets = info.sheets.filter((s) => !s.hidden || includeHidden).map((s) => s.name);
  const [sheets, setSheets] = useState<string[]>(override.sheets ?? defaultSheets);
  const [headerRow, setHeaderRow] = useState<Record<string, number>>(override.headerRow ?? {});
  const [encoding, setEncoding] = useState(override.encoding ?? '');
  const [delimiter, setDelimiter] = useState(override.delimiter ?? '');
  const [active, setActive] = useState(info.sheets.find((s) => sheets.includes(s.name))?.name ?? info.sheets[0]?.name);
  const sheet = info.sheets.find((s) => s.name === active);
  const firstRow = sheet?.firstRow ?? 1;
  const current = sheet ? headerRow[sheet.name] ?? sheet.headerRow : 1;

  const save = () => {
    const ov: FileReadOverride = {};
    const sameSheets = sheets.length === defaultSheets.length && sheets.every((s) => defaultSheets.includes(s));
    if (!sameSheets) ov.sheets = sheets;
    if (Object.keys(headerRow).length) ov.headerRow = headerRow;
    if (encoding) ov.encoding = encoding;
    if (delimiter) ov.delimiter = delimiter;
    onSave(ov);
  };

  return (
    <Modal title={`Settings — ${info.name}`} onClose={onClose} footer={<>
      <button className="btn" onClick={onClose}>Cancel</button>
      <button className="btn primary" disabled={!sheets.length} onClick={save}>Save and re-check</button>
    </>}>
      {isText && (
        <div className="grid2">
          <Field label={`Text encoding (detected: ${info.encoding})`}>
            <select value={encoding} onChange={(e) => setEncoding(e.target.value)}>
              <option value="">Automatic</option>
              {ENCODINGS.map((e) => <option key={e.id} value={e.id}>{e.label}</option>)}
            </select>
          </Field>
          <Field label={`Delimiter (detected: ${info.delimiter === '\t' ? 'Tab' : info.delimiter})`}>
            <select value={delimiter} onChange={(e) => setDelimiter(e.target.value)}>
              <option value="">Automatic</option>
              {DELIMITERS.map((d) => <option key={d.id} value={d.id}>{d.label}</option>)}
            </select>
          </Field>
        </div>
      )}
      {!isText && (
        <>
          <div className="field-label">Sheets to process</div>
          <div className="list" style={{ marginBottom: 14 }}>
            {info.sheets.map((s) => (
              <div className="list-item" key={s.name}>
                <input type="checkbox" checked={sheets.includes(s.name)} onChange={(e) => setSheets((cur) => (e.target.checked ? [...cur, s.name] : cur.filter((x) => x !== s.name)))} aria-label={`Include ${s.name}`} />
                <button className="btn link name" style={{ textAlign: 'left' }} onClick={() => setActive(s.name)}>{s.name}</button>
                {s.hidden && <Tag tone="warn">Hidden</Tag>}
                {s.hasFormulas && <Tag>Formulas</Tag>}
                {s.mergeCount > 0 && <Tag>{s.mergeCount} merged</Tag>}
                <span className="meta">{s.rowCount.toLocaleString('en-IN')} rows × {s.colCount} cols</span>
                {active === s.name && <Tag tone="busy">Showing</Tag>}
              </div>
            ))}
          </div>
        </>
      )}
      {sheet && (
        <>
          <div className="row">
            <div className="field-label grow">Header row for “{sheet.name}” — click the row that holds the column names</div>
            {headerRow[sheet.name] && <button className="btn small" onClick={() => setHeaderRow((h) => { const n = { ...h }; delete n[sheet.name]; return n; })}>Use automatic (row {sheet.detectedHeaderRow})</button>}
          </div>
          <DataPreview
            rows={sheet.preview}
            rowNumbers={sheet.preview.map((_, i) => firstRow + i)}
            chosenRow={current - firstRow}
            onPickRow={(i) => setHeaderRow((h) => ({ ...h, [sheet.name]: firstRow + i }))}
          />
          <p className="muted small">Header row: {current}{headerRow[sheet.name] ? ' (chosen)' : ' (detected)'}. Rows above it are kept as title rows when the layout is unchanged. Only the first 15 rows are shown here.</p>
          {current - firstRow >= sheet.preview.length && <Note tone="warn">The header row is below the rows shown.</Note>}
        </>
      )}
      {!isText && <p className="muted small">Password-protected (encrypted) workbooks cannot be opened here; remove the password in Excel first.</p>}
    </Modal>
  );
}
