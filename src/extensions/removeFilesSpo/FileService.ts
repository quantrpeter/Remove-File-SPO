import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * The criteria the user picks in the dialog.
 * Either remove files larger than a size, or remove files older than a number of weeks.
 */
export type RemoveMode = 'size' | 'age';

export interface IRemoveCriteria {
  mode: RemoveMode;
  /** Minimum size in bytes (exclusive). Used when mode === 'size'. */
  minSizeBytes?: number;
  /** Files last modified more than this many weeks ago. Used when mode === 'age'. */
  olderThanWeeks?: number;
}

/** A file in the current library that is a candidate for removal. */
export interface IRemovableFile {
  id: number;
  name: string;
  serverRelativeUrl: string;
  sizeBytes: number;
  modified: string;
}

/** Everything FileService needs to talk to the current library. */
export interface IFileQueryContext {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  listId: string;
  /** Server-relative URL of the library root, e.g. "/sites/backup/Backup". */
  listServerRelativeUrl: string;
  /** Display name of the library, used as the first breadcrumb level. */
  listTitle: string;
  currentUserId: number;
}

export interface IRecycleResult {
  id: number;
  ok: boolean;
  error?: string;
}

interface ISpItem {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Modified: string;
  File?: { Length: string };
}

/**
 * Loads the current user's files in the library, filtered by the chosen criteria.
 * Author filtering and the file/folder split happen server side; size and age are
 * applied client side because they are not reliably filterable in OData.
 */
export async function getMyFiles(
  ctx: IFileQueryContext,
  criteria: IRemoveCriteria
): Promise<IRemovableFile[]> {
  const select: string =
    '$select=Id,FileLeafRef,FileRef,Modified,File/Length&$expand=File';
  const filter: string = `$filter=AuthorId eq ${ctx.currentUserId} and FSObjType eq 0`;
  const url: string =
    `${ctx.webAbsoluteUrl}/_api/web/lists(guid'${ctx.listId}')/items` +
    `?${select}&${filter}&$top=5000`;

  const response: SPHttpClientResponse = await ctx.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    const text: string = await response.text();
    throw new Error(`Failed to load files (HTTP ${response.status}). ${text}`);
  }

  const json: { value: ISpItem[] } = await response.json();
  let files: IRemovableFile[] = (json.value || []).map((it: ISpItem) => ({
    id: it.Id,
    name: it.FileLeafRef,
    serverRelativeUrl: it.FileRef,
    sizeBytes: it.File ? Number(it.File.Length) : 0,
    modified: it.Modified
  }));

  if (criteria.mode === 'size' && criteria.minSizeBytes !== undefined) {
    const min: number = criteria.minSizeBytes;
    files = files.filter((f: IRemovableFile) => f.sizeBytes > min);
  } else if (criteria.mode === 'age' && criteria.olderThanWeeks !== undefined) {
    const cutoff: number =
      Date.now() - criteria.olderThanWeeks * 7 * 24 * 60 * 60 * 1000;
    files = files.filter(
      (f: IRemovableFile) => new Date(f.modified).getTime() < cutoff
    );
  }

  files.sort((a: IRemovableFile, b: IRemovableFile) => b.sizeBytes - a.sizeBytes);
  return files;
}

/** Called after each file is processed so callers can show progress. */
export type RecycleProgress = (completed: number, total: number) => void;

/**
 * Sends the given items to the site recycle bin (safer than a permanent delete).
 * Requests run sequentially to avoid throttling on larger batches.
 * Invokes onProgress after each item so the caller can render progress.
 */
export async function recycleFiles(
  ctx: IFileQueryContext,
  ids: number[],
  onProgress?: RecycleProgress
): Promise<IRecycleResult[]> {
  const results: IRecycleResult[] = [];
  const total: number = ids.length;

  for (const id of ids) {
    const url: string =
      `${ctx.webAbsoluteUrl}/_api/web/lists(guid'${ctx.listId}')/items(${id})/recycle()`;
    try {
      const response: SPHttpClientResponse = await ctx.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        { headers: { 'IF-MATCH': '*' } }
      );
      results.push({
        id,
        ok: response.ok,
        error: response.ok ? undefined : `HTTP ${response.status}`
      });
    } catch (error) {
      results.push({ id, ok: false, error: String(error) });
    }
    if (onProgress) {
      onProgress(results.length, total);
    }
  }

  return results;
}

/** Human readable byte size, e.g. 1536 -> "1.5 KB". */
export function formatBytes(bytes: number): string {
  if (!bytes || bytes < 0) {
    return '0 B';
  }
  const units: string[] = ['B', 'KB', 'MB', 'GB', 'TB'];
  const exponent: number = Math.min(
    Math.floor(Math.log(bytes) / Math.log(1024)),
    units.length - 1
  );
  const value: number = bytes / Math.pow(1024, exponent);
  return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}
