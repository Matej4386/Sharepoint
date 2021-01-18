import { FilterTemplateOption } from './FilterTemplateOption';
export interface IFilterVal {
    title: string;
    count: number;
}
export interface IFilters {
    /**
     * Display name for filter
     */
    filterTitle: string;
    /**
     * Type of column
     */
    filterType: string;
    /**
     * Internal filter name for column
     */
    filterName: string;
    /**
     * Filter values unique
     */
    filterValues: IFilterVal[];
    /**
     * Mode for filter - defualt, checkbox, multi value, date ...
     */
    filterMode: FilterTemplateOption;
    /**
     * For multi value filters
     */
    executeSearch: boolean;
    /**
     * Allow refiners to be expanded by default
     */
    showExpanded: boolean;
    /**
     * Show filter textbox to search inside the refiner values
     */
    showValueFilter: boolean;
}