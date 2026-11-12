/// <reference lib="webworker" />
// All reading, processing and writing happens in this worker so the page stays responsive.
// The worker is stateless: every request carries the File handles it needs, so "Cancel" can simply
// terminate the worker and start a fresh one without leaving anything half-written.

import { inspectFile } from '../core/parse';
import { previewBatch, runBatch } from '../core/pipeline';
import { ProcessError } from '../core/ops/helpers';
import type { WorkerRequest, WorkerResponse } from './protocol';

const ctx = self as unknown as DedicatedWorkerGlobalScope;
const post = (msg: WorkerResponse, transfer: Transferable[] = []) => ctx.postMessage(msg, transfer);

ctx.onmessage = async (ev: MessageEvent<WorkerRequest>) => {
  const req = ev.data;
  try {
    if (req.type === 'inspect') {
      const infos = [];
      for (let i = 0; i < req.files.length; i++) {
        const f = req.files[i];
        post({ type: 'progress', reqId: req.reqId, event: { phase: 'read', pct: Math.round((i / req.files.length) * 100), message: `Checking ${f.name}`, fileId: f.id } });
        infos.push((await inspectFile(f, req.overrides[f.id])).info);
      }
      post({ type: 'inspected', reqId: req.reqId, infos });
      return;
    }
    const report = (event: Parameters<Parameters<typeof previewBatch>[3]>[0]) => post({ type: 'progress', reqId: req.reqId, event });
    if (req.type === 'preview') {
      const result = await previewBatch(req.files, req.infos, req.config, report);
      post({ type: 'previewed', reqId: req.reqId, result });
      return;
    }
    if (req.type === 'run') {
      const result = await runBatch(req.files, req.infos, req.config, report);
      const buffers = new Set<ArrayBuffer>([result.zip.buffer as ArrayBuffer]);
      for (const f of result.files) buffers.add(f.bytes.buffer as ArrayBuffer);
      post({ type: 'done', reqId: req.reqId, result }, [...buffers]);
    }
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    post({ type: 'failed', reqId: req.reqId, message, suggestion: e instanceof ProcessError ? e.suggestion : undefined });
  }
};
