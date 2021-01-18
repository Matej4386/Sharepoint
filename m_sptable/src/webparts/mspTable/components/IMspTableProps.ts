import {IContextualMenuItem} from 'office-ui-fabric-react/lib/ContextualMenu';
import {IFilterConfiguration} from './modules/datatypes/IFilterConfiguration';
import {SelectionMode} from 'office-ui-fabric-react/lib/DetailsList';

export interface IMspTableProps {
    /**
     * from context not to pass whole context if it is possible
     */
    currentCultureName: string;
    /**
     * console.log info
     */
    debug: boolean;
    /**
     * List ID to display
     */
    listToDisplayId: string;
    /**
     * Filters configuration
     */
    filterConfiguration: IFilterConfiguration[];
    /**
     * String with internal column names separated with ,
     */
    searchConfiguration: string;
    /**
     * Initial sorting Internal Column Name
     */
    onInitSortedKey: string;
    /**
     * selection mode - none, single, multiple
     */
    selectionMode: SelectionMode;
    /**
     * render own commandbar?
     */
    commandBar?: boolean;
    /**
     * if we have parent command bar this will send new far command bar items after read data
     */
    onCommandBarChange?: (newMainCommands: IContextualMenuItem[], farItems: IContextualMenuItem[]) => void;
    /**
     * function to render info (errors, success) in parent element
     */
    renderInfo: (error: boolean, message: string) => void;
}
