import {IFilters} from '../../../datatypes/IFilters';
export interface IMCollapseProps {
    filter: IFilters;
    items: JSX.Element[];
    groupIndex: number;
    isSelectedItem: boolean;
    filterField: string;
    onFilter: (filter: IFilters, filterValue: string[], add: boolean, multi: boolean, executeSearch: boolean) => void;
}