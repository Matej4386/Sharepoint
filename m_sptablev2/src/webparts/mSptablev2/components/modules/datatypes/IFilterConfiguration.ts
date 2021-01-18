import {FilterTemplateOption} from './FilterTemplateOption';
import {FiltersSortDirection} from './FiltersSortDirection';
import {FiltersSortOption}  from './FiltersSortOptions';
export interface IFilterConfiguration {
    /**
     * The SharePoint refiner name
     */
    filterName: string;

    /**
     * The selected template for this refiner
     */
    filterMode: FilterTemplateOption;

    /**
     * How the refiner values should be sorted
     */
    filterSortType: FiltersSortOption;

    /**
     * Direction of sorting values
     */
    filterSortDirection: FiltersSortDirection;

    /**
     * Allow refiners to be expanded by default
     */
    showExpanded: boolean;

    /**
     * Show filter textbox to search inside the refiner values
     */
    showValueFilter: boolean;
    /**
     * Show filter in default filter view
     */
    inDefaultFilterView: boolean;
}