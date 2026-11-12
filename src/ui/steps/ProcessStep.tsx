import { useEffect, useRef } from 'react';
import type { FileInfo } from '../../core/types';
import type { ProgressEvent } from '../../core/pipeline';
import { formatBytes } from '../../core/util';
import { Note, PageHead, Progress, Tag } from '../components/common';

export type FileRunState = 'queued' | 'reading' | 'processing' | 'exporting' | 'done' | 'review' | 'error' | 'skipped';

const LABEL: Record<FileRunState, [string, 'ok' | 'warn' | 'bad' | 'busy' | undefined]> = {
  queued: ['Waiting', undefined], reading: ['Reading', 'busy'], processing: ['Processing', 'busy'], exporting: ['Writing', 'busy'],
  done: ['Done', 'ok'], review: ['Needs review', 'warn'], error: ['Error', 'bad'], skipped: ['Skipped', undefined],
};

interface Props {
  infos: FileInfo[];
  running: boolean;
  progress: ProgressEvent | null;
  states: Record<string, FileRunState>;
  log: string[];
  error: { message: string; suggestion?: string } | null;
  onStart: () => void;
  onCancel: () => void;
  done: boolean;
}

export function ProcessStep({ infos, running, progress, states, log, error, onStart, onCancel, done }: Props) {
  const logRef = useRef<HTMLDivElement>(null);
  useEffect(() => { logRef.current?.scrollTo({ top: logRef.current.scrollHeight }); }, [log]);
  const started = running || done || !!error || log.length > 0;
  return (
    <>
      <PageHead title={running ? 'Processing…' : done ? 'Processing finished' : 'Ready to process'} sub="Everything runs inside this browser tab. Keep the tab open until it finishes." />
      {!started && (
        <div className="card">
          <p>Your files will be processed and packed into a ZIP. Originals are not changed.</p>
          <button className="btn primary" onClick={onStart}>Start processing</button>
        </div>
      )}
      {started && (
        <div className="card">
          <div className="row" style={{ marginBottom: 8 }}>
            <b className="grow">{progress?.message ?? ''}</b>
            <span className="muted small">{progress?.pct ?? 0}%</span>
            {running && <button className="btn small danger" onClick={onCancel}>Stop</button>}
          </div>
          <Progress pct={progress?.pct ?? 0} />
        </div>
      )}
      {error && <Note tone="bad">{error.message}{error.suggestion && <div>{error.suggestion}</div>}<div style={{ marginTop: 8 }}><button className="btn small" onClick={onStart}>Try again</button></div></Note>}
      {started && (
        <div className="card">
          <h2>Files</h2>
          <div className="list" style={{ marginTop: 8 }}>
            {infos.map((i) => {
              const st = states[i.id] ?? 'queued';
              const [label, tone] = LABEL[st];
              return (
                <div className="list-item" key={i.id}>
                  <span className="name" title={i.relPath}>{i.relPath}</span>
                  <span className="meta">{formatBytes(i.size)}</span>
                  <Tag tone={tone}>{label}</Tag>
                </div>
              );
            })}
          </div>
        </div>
      )}
      {log.length > 0 && (
        <div className="card">
          <h2>Activity</h2>
          <div className="log" ref={logRef} aria-live="polite">{log.join('\n')}</div>
        </div>
      )}
    </>
  );
}
