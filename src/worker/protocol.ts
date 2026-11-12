import type { FileInfo, SourceFile } from '../core/types';
import type { FileReadOverride, WorkflowConfig } from '../core/workflow';
import type { PreviewResult, ProgressEvent, RunResult } from '../core/pipeline';

export type WorkerRequest =
  | { type: 'inspect'; reqId: number; files: SourceFile[]; overrides: Record<string, FileReadOverride> }
  | { type: 'preview'; reqId: number; files: SourceFile[]; infos: FileInfo[]; config: WorkflowConfig }
  | { type: 'run'; reqId: number; files: SourceFile[]; infos: FileInfo[]; config: WorkflowConfig };

export type WorkerResponse =
  | { type: 'progress'; reqId: number; event: ProgressEvent }
  | { type: 'inspected'; reqId: number; infos: FileInfo[] }
  | { type: 'previewed'; reqId: number; result: PreviewResult }
  | { type: 'done'; reqId: number; result: RunResult }
  | { type: 'failed'; reqId: number; message: string; suggestion?: string };
