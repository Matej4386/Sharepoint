import {IContextualMenuProps} from 'office-ui-fabric-react/lib/ContextualMenu';
import {IColumn} from 'office-ui-fabric-react/lib/DetailsList';
import {IContextualMenuItem} from 'office-ui-fabric-react/lib/ContextualMenu';
export interface IMSptablev2State {
    items: any;
    currentViewId: string;
    columnsDefinition: any[];
    columns: IColumn[];
    isLoadingData: boolean;
    contextualMenuProps?: IContextualMenuProps;
    sortedColumnKey?: string;
    isSortedDescending?: boolean;
    selectionDetails: any[];
    commandBarItems: IContextualMenuItem[];
    farItems: IContextualMenuItem[];
    nrPreloadingIterations: number;
    /**
     * Filters
     */
    showFiltersPanel: boolean;
    filterField: string;
}