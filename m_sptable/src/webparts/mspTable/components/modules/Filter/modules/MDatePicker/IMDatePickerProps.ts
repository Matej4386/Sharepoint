import {IFilters} from '../../../datatypes/IFilters';
export interface IMDatePickerProps {
    currentCultureName: string;
    filter: IFilters;
    isSelectedItem: boolean;
    onFilter: (filter: IFilters, filterValue: string[], add: boolean, multi: boolean, executeSearch: boolean) => void;
}