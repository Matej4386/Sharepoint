import {IFilters} from '../datatypes/IFilters';
export interface IFilterPanelProps {
    currentCultureName: string;
    debug: boolean;
    showPanel: boolean;
    filters: IFilters[];
    selectedFilters: IFilters[];
    filterField: string;
    renderInfo: (error: boolean, message: string) => void;
    onUpdateShow: (show: boolean) => void;
    onFilter: (filter: IFilters, filterValue: string[], add: boolean, multi: boolean, executeSearch: boolean) => void;
    onRemoveAllFilters: () => void;
}