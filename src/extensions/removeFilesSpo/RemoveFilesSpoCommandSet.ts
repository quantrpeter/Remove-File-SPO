import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import RemoveFilesDialog from './RemoveFilesDialog';
import { type IFileQueryContext } from './FileService';

/**
 * This command set has no configurable client-side properties.
 */
export type IRemoveFilesSpoCommandSetProperties = Record<string, unknown>;

const LOG_SOURCE: string = 'RemoveFilesSpoCommandSet';

export default class RemoveFilesSpoCommandSet extends BaseListViewCommandSet<IRemoveFilesSpoCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized RemoveFilesSpoCommandSet');
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'REMOVE_FILES':
        this._openRemoveFilesDialog();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _openRemoveFilesDialog(): void {
    const list = this.context.pageContext.list;
    if (!list) {
      Dialog.alert('This command can only be used inside a document library.').catch(() => {
        /* ignore */
      });
      return;
    }

    const queryContext: IFileQueryContext = {
      spHttpClient: this.context.spHttpClient,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      listId: list.id.toString(),
      listServerRelativeUrl: list.serverRelativeUrl,
      listTitle: list.title,
      currentUserId: this.context.pageContext.legacyPageContext.userId
    };

    const dialog: RemoveFilesDialog = new RemoveFilesDialog(queryContext);
    dialog
      .show()
      .then(() => {
        // Refresh the list view so removed files disappear from the page.
        if (dialog.deletedCount > 0) {
          window.location.reload();
        }
      })
      .catch((error: Error) => {
        Log.error(LOG_SOURCE, error);
      });
  }
}
