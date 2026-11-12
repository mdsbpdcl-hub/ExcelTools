import type { WorkerRequest, WorkerResponse } from '../worker/protocol';
import type { ProgressEvent } from '../core/pipeline';

type Pending = { resolve: (v: WorkerResponse) => void; reject: (e: Error) => void; onProgress?: (e: ProgressEvent) => void };

export class CancelledError extends Error {
  constructor() { super('Cancelled'); }
}

export class TaskFailed extends Error {
  constructor(message: string, public suggestion?: string) { super(message); }
}

type Body<T> = T extends unknown ? Omit<T, 'reqId'> : never;

/** Talks to the processing worker. cancel() stops the current job immediately by restarting the worker. */
export class WorkerClient {
  private worker!: Worker;
  private nextId = 1;
  private pending = new Map<number, Pending>();

  constructor() { this.start(); }

  private start() {
    this.worker = new Worker(new URL('../worker/processor.worker.ts', import.meta.url), { type: 'module' });
    this.worker.onmessage = (ev: MessageEvent<WorkerResponse>) => {
      const msg = ev.data;
      const p = this.pending.get(msg.reqId);
      if (!p) return;
      if (msg.type === 'progress') { p.onProgress?.(msg.event); return; }
      this.pending.delete(msg.reqId);
      if (msg.type === 'failed') p.reject(new TaskFailed(msg.message, msg.suggestion));
      else p.resolve(msg);
    };
    this.worker.onerror = (ev) => {
      const err = new TaskFailed(`The processing engine stopped unexpectedly${ev.message ? `: ${ev.message}` : ''}. This usually means the browser ran out of memory.`, 'Process fewer or smaller files at a time, or close other tabs.');
      for (const p of this.pending.values()) p.reject(err);
      this.pending.clear();
      this.worker.terminate();
      this.start();
    };
  }

  request<T extends WorkerResponse>(body: Body<WorkerRequest>, onProgress?: (e: ProgressEvent) => void): Promise<T> {
    const reqId = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      this.pending.set(reqId, { resolve: resolve as (v: WorkerResponse) => void, reject, onProgress });
      this.worker.postMessage({ ...body, reqId } as WorkerRequest);
    });
  }

  cancel(): void {
    this.worker.terminate();
    for (const p of this.pending.values()) p.reject(new CancelledError());
    this.pending.clear();
    this.start();
  }
}
