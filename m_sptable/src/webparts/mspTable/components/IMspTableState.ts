import {IContextualMenuProps} from 'office-ui-fabric-react/lib/ContextualMenu';
import {IColumn} from 'office-ui-fabric-react/lib/DetailsList';
import {IFilters}  from './modules/datatypes/IFilters';
import {IContextualMenuItem} from 'office-ui-fabric-react/lib/ContextualMenu';
export interface IMspTableState {
    items: any;
    columnsDefinition: any;
    columns: IColumn[];
    isLoadingData: boolean;
    contextualMenuProps?: IContextualMenuProps;
    sortedColumnKey?: string;
    isSortedDescending?: boolean;
    filters: IFilters[];
    selectedFilters: IFilters[];
    showFiltersPanel: boolean;
    filterField: string;
    valueFilter: string;
    editValueText: string;
    selectionDetails: any[];
    commandBarItems: IContextualMenuItem[];
    farItems: IContextualMenuItem[];
}