import { BaseDialog } from '@microsoft/sp-dialog';
import {
  getMyFiles,
  recycleFiles,
  formatBytes,
  type IFileQueryContext,
  type IRemovableFile,
  type IRemoveCriteria,
  type RemoveMode
} from './FileService';

type DialogStep = 'criteria' | 'loading' | 'list' | 'deleting' | 'done' | 'error';

interface IDeleteProgress {
  completed: number;
  total: number;
}

const STYLE_ELEMENT_ID: string = 'removeFilesSpoDialogStyles';

/**
 * Self contained dialog that walks the user through:
 *   1. choosing removal criteria (size larger than X, or older than N weeks)
 *   2. reviewing the matching files and confirming the removal
 * Removed files are sent to the recycle bin.
 */
export default class RemoveFilesDialog extends BaseDialog {
  /** Number of files actually recycled. Read by the caller to decide whether to refresh. */
  public deletedCount: number = 0;

  private _ctx: IFileQueryContext;
  private _step: DialogStep = 'criteria';
  private _mode: RemoveMode = 'size';
  private _sizeValueMb: number = 100;
  private _weeksValue: number = 4;
  private _files: IRemovableFile[] = [];
  private _selectedIds: Set<number> = new Set<number>();
  private _errorMessage: string = '';
  private _deleteProgress: IDeleteProgress = { completed: 0, total: 0 };

  constructor(ctx: IFileQueryContext) {
    super({ isBlocking: false });
    this._ctx = ctx;
  }

  protected render(): void {
    this._injectStyles();

    const root: HTMLElement = this.domElement;
    root.innerHTML = '';

    const container: HTMLDivElement = document.createElement('div');
    container.className = 'rfs-dialog';
    if (this._step === 'list' && this._files.length > 0) {
      container.className += ' rfs-dialog--wide';
    }

    switch (this._step) {
      case 'criteria':
        this._renderCriteria(container);
        break;
      case 'loading':
        this._renderMessage(container, 'Finding files\u2026', true);
        break;
      case 'list':
        this._renderList(container);
        break;
      case 'deleting':
        this._renderProgress(container);
        break;
      case 'done':
        this._renderDone(container);
        break;
      case 'error':
        this._renderError(container);
        break;
      default:
        break;
    }

    root.appendChild(container);
  }

  // ---- Step 1: criteria selection -----------------------------------------

  private _renderCriteria(container: HTMLDivElement): void {
    container.appendChild(this._header('Remove Files'));

    const intro: HTMLParagraphElement = document.createElement('p');
    intro.className = 'rfs-text';
    intro.textContent = 'Choose which of your files to remove from this library.';
    container.appendChild(intro);

    const sizeRow: HTMLDivElement = this._radioOption(
      'rfs-mode',
      this._mode === 'size',
      'Files larger than',
      'MB',
      String(this._sizeValueMb),
      (checked: boolean) => {
        if (checked) {
          this._mode = 'size';
          this.render();
        }
      },
      (value: string) => {
        this._sizeValueMb = Math.max(0, Number(value) || 0);
      },
      this._mode === 'size'
    );
    container.appendChild(sizeRow);

    const ageRow: HTMLDivElement = this._radioOption(
      'rfs-mode',
      this._mode === 'age',
      'Files older than',
      'weeks',
      String(this._weeksValue),
      (checked: boolean) => {
        if (checked) {
          this._mode = 'age';
          this.render();
        }
      },
      (value: string) => {
        this._weeksValue = Math.max(0, Number(value) || 0);
      },
      this._mode === 'age'
    );
    container.appendChild(ageRow);

    const footer: HTMLDivElement = document.createElement('div');
    footer.className = 'rfs-footer';
    footer.appendChild(
      this._button('Cancel', 'rfs-btn rfs-btn--secondary', () => {
        this.close().catch(() => { /* ignore */ });
      })
    );
    footer.appendChild(
      this._button('Find files', 'rfs-btn rfs-btn--primary', () => {
        this._findFiles().catch((e: Error) => this._fail(e));
      })
    );
    container.appendChild(footer);
  }

  private async _findFiles(): Promise<void> {
    this._step = 'loading';
    this.render();

    const criteria: IRemoveCriteria =
      this._mode === 'size'
        ? { mode: 'size', minSizeBytes: this._sizeValueMb * 1024 * 1024 }
        : { mode: 'age', olderThanWeeks: this._weeksValue };

    this._files = await getMyFiles(this._ctx, criteria);
    this._selectedIds = new Set<number>(this._files.map((f: IRemovableFile) => f.id));
    this._step = 'list';
    this.render();
  }

  // ---- Step 2: review and confirm -----------------------------------------

  private _renderList(container: HTMLDivElement): void {
    container.appendChild(this._header('Confirm removal'));

    if (this._files.length === 0) {
      const empty: HTMLParagraphElement = document.createElement('p');
      empty.className = 'rfs-text';
      empty.textContent = 'No files matched your criteria.';
      container.appendChild(empty);

      const footer: HTMLDivElement = document.createElement('div');
      footer.className = 'rfs-footer';
      footer.appendChild(
        this._button('Back', 'rfs-btn rfs-btn--secondary', () => {
          this._step = 'criteria';
          this.render();
        })
      );
      footer.appendChild(
        this._button('Close', 'rfs-btn rfs-btn--primary', () => {
          this.close().catch(() => { /* ignore */ });
        })
      );
      container.appendChild(footer);
      return;
    }

    const summary: HTMLParagraphElement = document.createElement('p');
    summary.className = 'rfs-text';
    summary.textContent =
      `${this._files.length} file(s) matched. Review the list and confirm what to remove (they will go to the recycle bin).`;
    container.appendChild(summary);

    const table: HTMLDivElement = document.createElement('div');
    table.className = 'rfs-table';

    for (const file of this._files) {
      const row: HTMLLabelElement = document.createElement('label');
      row.className = 'rfs-row';

      const checkbox: HTMLInputElement = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = this._selectedIds.has(file.id);
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          this._selectedIds.add(file.id);
        } else {
          this._selectedIds.delete(file.id);
        }
        this._updateDeleteButton();
      });

      const path: HTMLDivElement = this._buildBreadcrumb(file.serverRelativeUrl);
      path.classList.add('rfs-row__path');

      const main: HTMLDivElement = document.createElement('div');
      main.className = 'rfs-row__main';

      const name: HTMLAnchorElement = document.createElement('a');
      name.className = 'rfs-row__name';
      name.textContent = file.name;
      name.href = this._absoluteUrl(file.serverRelativeUrl);
      name.target = '_blank';
      name.rel = 'noopener noreferrer';
      name.title = file.serverRelativeUrl;
      // Don't toggle the checkbox when the user clicks the file link.
      name.addEventListener('click', (e: MouseEvent) => e.stopPropagation());

      main.appendChild(name);

      const meta: HTMLSpanElement = document.createElement('span');
      meta.className = 'rfs-row__meta';
      meta.textContent = `${formatBytes(file.sizeBytes)} \u00b7 ${this._formatDate(file.modified)}`;

      row.appendChild(checkbox);
      row.appendChild(path);
      row.appendChild(main);
      row.appendChild(meta);
      table.appendChild(row);
    }

    container.appendChild(table);

    const footer: HTMLDivElement = document.createElement('div');
    footer.className = 'rfs-footer';
    footer.appendChild(
      this._button('Back', 'rfs-btn rfs-btn--secondary', () => {
        this._step = 'criteria';
        this.render();
      })
    );
    const deleteButton: HTMLButtonElement = this._button(
      this._deleteLabel(),
      'rfs-btn rfs-btn--danger',
      () => {
        this._deleteSelected().catch((e: Error) => this._fail(e));
      }
    );
    deleteButton.setAttribute('data-role', 'delete');
    deleteButton.disabled = this._selectedIds.size === 0;
    footer.appendChild(deleteButton);
    container.appendChild(footer);
  }

  private _updateDeleteButton(): void {
    const button: HTMLButtonElement | null = this.domElement.querySelector(
      'button[data-role="delete"]'
    );
    if (button) {
      button.textContent = this._deleteLabel();
      button.disabled = this._selectedIds.size === 0;
    }
  }

  private _deleteLabel(): string {
    return `Remove ${this._selectedIds.size} file(s)`;
  }

  private async _deleteSelected(): Promise<void> {
    const ids: number[] = Array.from(this._selectedIds);
    if (ids.length === 0) {
      return;
    }
    this._step = 'deleting';
    this._deleteProgress = { completed: 0, total: ids.length };
    this.render();

    const results = await recycleFiles(
      this._ctx,
      ids,
      (completed: number, total: number) => {
        this._deleteProgress = { completed, total };
        this._updateProgress();
      }
    );
    this.deletedCount = results.filter((r) => r.ok).length;
    const failed: number = results.length - this.deletedCount;
    this._errorMessage =
      failed > 0 ? `${failed} file(s) could not be removed.` : '';
    this._step = 'done';
    this.render();
  }

  // ---- Final + helper screens ---------------------------------------------

  private _renderDone(container: HTMLDivElement): void {
    container.appendChild(this._header('Done'));

    const msg: HTMLParagraphElement = document.createElement('p');
    msg.className = 'rfs-text';
    msg.textContent = `${this.deletedCount} file(s) moved to the recycle bin.`;
    container.appendChild(msg);

    if (this._errorMessage) {
      const err: HTMLParagraphElement = document.createElement('p');
      err.className = 'rfs-text rfs-text--error';
      err.textContent = this._errorMessage;
      container.appendChild(err);
    }

    const footer: HTMLDivElement = document.createElement('div');
    footer.className = 'rfs-footer';
    footer.appendChild(
      this._button('Close', 'rfs-btn rfs-btn--primary', () => {
        this.close().catch(() => { /* ignore */ });
      })
    );
    container.appendChild(footer);
  }

  private _renderError(container: HTMLDivElement): void {
    container.appendChild(this._header('Something went wrong'));

    const err: HTMLParagraphElement = document.createElement('p');
    err.className = 'rfs-text rfs-text--error';
    err.textContent = this._errorMessage;
    container.appendChild(err);

    const footer: HTMLDivElement = document.createElement('div');
    footer.className = 'rfs-footer';
    footer.appendChild(
      this._button('Back', 'rfs-btn rfs-btn--secondary', () => {
        this._step = 'criteria';
        this.render();
      })
    );
    footer.appendChild(
      this._button('Close', 'rfs-btn rfs-btn--primary', () => {
        this.close().catch(() => { /* ignore */ });
      })
    );
    container.appendChild(footer);
  }

  private _renderMessage(
    container: HTMLDivElement,
    message: string,
    spinner: boolean
  ): void {
    const wrapper: HTMLDivElement = document.createElement('div');
    wrapper.className = 'rfs-loading';
    if (spinner) {
      const dot: HTMLDivElement = document.createElement('div');
      dot.className = 'rfs-spinner';
      wrapper.appendChild(dot);
    }
    const text: HTMLSpanElement = document.createElement('span');
    text.className = 'rfs-text';
    text.textContent = message;
    wrapper.appendChild(text);
    container.appendChild(wrapper);
  }

  private _renderProgress(container: HTMLDivElement): void {
    container.appendChild(this._header('Removing files'));

    const text: HTMLParagraphElement = document.createElement('p');
    text.className = 'rfs-text';
    text.setAttribute('data-role', 'progress-text');
    text.textContent = this._progressLabel();
    container.appendChild(text);

    const track: HTMLDivElement = document.createElement('div');
    track.className = 'rfs-progress';

    const fill: HTMLDivElement = document.createElement('div');
    fill.className = 'rfs-progress__fill';
    fill.setAttribute('data-role', 'progress-fill');
    fill.style.width = `${this._progressPercent()}%`;

    track.appendChild(fill);
    container.appendChild(track);
  }

  private _updateProgress(): void {
    const text: HTMLParagraphElement | null = this.domElement.querySelector(
      '[data-role="progress-text"]'
    );
    if (text) {
      text.textContent = this._progressLabel();
    }
    const fill: HTMLElement | null = this.domElement.querySelector(
      '[data-role="progress-fill"]'
    );
    if (fill) {
      fill.style.width = `${this._progressPercent()}%`;
    }
  }

  private _progressLabel(): string {
    const { completed, total } = this._deleteProgress;
    return `Removing ${completed} of ${total} file(s)\u2026`;
  }

  private _progressPercent(): number {
    const { completed, total } = this._deleteProgress;
    if (total === 0) {
      return 0;
    }
    return Math.round((completed / total) * 100);
  }

  private _fail(error: Error): void {
    this._errorMessage = error.message || String(error);
    this._step = 'error';
    this.render();
  }

  // ---- Small DOM builders --------------------------------------------------

  private _header(title: string): HTMLHeadingElement {
    const heading: HTMLHeadingElement = document.createElement('h2');
    heading.className = 'rfs-title';
    heading.textContent = title;
    return heading;
  }

  private _button(
    label: string,
    className: string,
    onClick: () => void
  ): HTMLButtonElement {
    const button: HTMLButtonElement = document.createElement('button');
    button.type = 'button';
    button.className = className;
    button.textContent = label;
    button.addEventListener('click', onClick);
    return button;
  }

  private _radioOption(
    radioName: string,
    checked: boolean,
    labelText: string,
    unit: string,
    inputValue: string,
    onToggle: (checked: boolean) => void,
    onInput: (value: string) => void,
    inputEnabled: boolean
  ): HTMLDivElement {
    const row: HTMLDivElement = document.createElement('div');
    row.className = 'rfs-option';

    const label: HTMLLabelElement = document.createElement('label');
    label.className = 'rfs-option__label';

    const radio: HTMLInputElement = document.createElement('input');
    radio.type = 'radio';
    radio.name = radioName;
    radio.checked = checked;
    radio.addEventListener('change', () => onToggle(radio.checked));

    const text: HTMLSpanElement = document.createElement('span');
    text.textContent = labelText;

    const numberInput: HTMLInputElement = document.createElement('input');
    numberInput.type = 'number';
    numberInput.min = '0';
    numberInput.className = 'rfs-number';
    numberInput.value = inputValue;
    numberInput.disabled = !inputEnabled;
    numberInput.addEventListener('input', () => onInput(numberInput.value));

    const unitText: HTMLSpanElement = document.createElement('span');
    unitText.className = 'rfs-unit';
    unitText.textContent = unit;

    label.appendChild(radio);
    label.appendChild(text);
    label.appendChild(numberInput);
    label.appendChild(unitText);
    row.appendChild(label);
    return row;
  }

  private _formatDate(iso: string): string {
    const date: Date = new Date(iso);
    if (isNaN(date.getTime())) {
      return iso;
    }
    return date.toLocaleDateString();
  }

  /** Site origin, e.g. "https://contoso.sharepoint.com". */
  private _origin(): string {
    try {
      return new URL(this._ctx.webAbsoluteUrl).origin;
    } catch {
      return '';
    }
  }

  /** Builds an absolute, properly encoded URL from a (decoded) server-relative path. */
  private _absoluteUrl(serverRelativeUrl: string): string {
    const encoded: string = serverRelativeUrl
      .split('/')
      .map((segment: string) => encodeURIComponent(segment))
      .join('/');
    return `${this._origin()}${encoded}`;
  }

  /** Link that opens a folder in the modern library view (new tab). */
  private _folderUrl(folderServerRelativeUrl: string): string {
    const libRoot: string = this._ctx.listServerRelativeUrl.replace(/\/+$/, '');
    return (
      `${this._origin()}${libRoot}/Forms/AllItems.aspx` +
      `?id=${encodeURIComponent(folderServerRelativeUrl)}`
    );
  }

  /**
   * Builds a clickable breadcrumb of every folder level for a file, starting at the
   * library root. Each level opens that folder in a new tab.
   */
  private _buildBreadcrumb(fileServerRelativeUrl: string): HTMLDivElement {
    const crumbs: HTMLDivElement = document.createElement('div');
    crumbs.className = 'rfs-crumbs';

    const libRoot: string = this._ctx.listServerRelativeUrl.replace(/\/+$/, '');

    crumbs.appendChild(
      this._crumbLink(this._ctx.listTitle || 'Library', this._folderUrl(libRoot))
    );

    const lastSlash: number = fileServerRelativeUrl.lastIndexOf('/');
    const folderPath: string =
      lastSlash > 0 ? fileServerRelativeUrl.substring(0, lastSlash) : '';

    let relative: string = folderPath;
    if (folderPath.toLowerCase().indexOf(libRoot.toLowerCase()) === 0) {
      relative = folderPath.substring(libRoot.length);
    }

    const segments: string[] = relative.split('/').filter((s: string) => s.length > 0);
    let cumulative: string = libRoot;
    for (const segment of segments) {
      cumulative += `/${segment}`;
      crumbs.appendChild(this._crumbSeparator());
      crumbs.appendChild(this._crumbLink(segment, this._folderUrl(cumulative)));
    }

    return crumbs;
  }

  private _crumbLink(label: string, href: string): HTMLAnchorElement {
    const link: HTMLAnchorElement = document.createElement('a');
    link.className = 'rfs-crumb';
    link.textContent = label;
    link.href = href;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.addEventListener('click', (e: MouseEvent) => e.stopPropagation());
    return link;
  }

  private _crumbSeparator(): HTMLSpanElement {
    const sep: HTMLSpanElement = document.createElement('span');
    sep.className = 'rfs-crumb-sep';
    sep.textContent = '/';
    return sep;
  }

  private _injectStyles(): void {
    if (document.getElementById(STYLE_ELEMENT_ID)) {
      return;
    }
    const style: HTMLStyleElement = document.createElement('style');
    style.id = STYLE_ELEMENT_ID;
    style.textContent = `
.rfs-dialog { font-family: 'Segoe UI', sans-serif; min-width: 480px; max-width: 920px; padding: 4px 4px 8px; color: #323130; }
.rfs-dialog--wide { width: 80vw; max-width: 920px; }
.rfs-title { font-size: 20px; font-weight: 600; margin: 0 0 12px; }
.rfs-text { font-size: 14px; line-height: 20px; margin: 0 0 12px; }
.rfs-text--error { color: #a4262c; }
.rfs-option { padding: 8px 0; border-top: 1px solid #edebe9; }
.rfs-option__label { display: flex; align-items: center; gap: 8px; font-size: 14px; }
.rfs-number { width: 80px; padding: 4px 6px; border: 1px solid #8a8886; border-radius: 2px; font-size: 14px; }
.rfs-number:disabled { background: #f3f2f1; color: #a19f9d; }
.rfs-unit { color: #605e5c; }
.rfs-table { max-height: 52vh; min-height: 200px; overflow-y: auto; border: 1px solid #edebe9; border-radius: 2px; margin-bottom: 16px; }
.rfs-row { display: grid; grid-template-columns: 24px minmax(0, 1fr) minmax(0, 1.4fr) auto; align-items: start; gap: 8px; padding: 8px 12px; font-size: 13px; border-bottom: 1px solid #f3f2f1; cursor: pointer; }
.rfs-row:last-child { border-bottom: none; }
.rfs-row:hover { background: #f3f2f1; }
.rfs-row input[type="checkbox"] { margin-top: 2px; }
.rfs-row__main { min-width: 0; }
.rfs-row__name { display: block; color: #0078d4; text-decoration: none; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rfs-row__name:hover { text-decoration: underline; }
.rfs-row__meta { color: #605e5c; white-space: nowrap; padding-top: 1px; }
.rfs-row__path { margin-top: 0; }
.rfs-crumbs { display: flex; flex-wrap: wrap; align-items: center; gap: 3px; margin-top: 3px; font-size: 11px; line-height: 16px; }
.rfs-crumb { color: #605e5c; text-decoration: none; }
.rfs-crumb:hover { color: #0078d4; text-decoration: underline; }
.rfs-crumb-sep { color: #a19f9d; }
.rfs-footer { display: flex; justify-content: flex-end; gap: 8px; margin-top: 8px; }
.rfs-btn { border: 1px solid transparent; border-radius: 2px; padding: 6px 16px; font-size: 14px; cursor: pointer; }
.rfs-btn:disabled { opacity: 0.5; cursor: default; }
.rfs-btn--primary { background: #0078d4; color: #fff; }
.rfs-btn--primary:hover:not(:disabled) { background: #106ebe; }
.rfs-btn--secondary { background: #fff; color: #323130; border-color: #8a8886; }
.rfs-btn--secondary:hover { background: #f3f2f1; }
.rfs-btn--danger { background: #a4262c; color: #fff; }
.rfs-btn--danger:hover:not(:disabled) { background: #8a1f24; }
.rfs-progress { height: 8px; background: #edebe9; border-radius: 4px; overflow: hidden; margin: 4px 0 16px; }
.rfs-progress__fill { height: 100%; width: 0; background: #0078d4; border-radius: 4px; transition: width 0.2s ease; }
.rfs-loading { display: flex; align-items: center; gap: 12px; padding: 24px 8px; }
.rfs-spinner { width: 20px; height: 20px; border: 2px solid #c8c6c4; border-top-color: #0078d4; border-radius: 50%; animation: rfs-spin 0.8s linear infinite; }
@keyframes rfs-spin { to { transform: rotate(360deg); } }
`;
    document.head.appendChild(style);
  }
}
