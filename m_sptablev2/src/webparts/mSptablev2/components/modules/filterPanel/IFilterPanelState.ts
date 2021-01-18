import {IFilter} from '../datatypes/IFilter';

export interface IFilterPanelState {
    /**
     * JSX Elements to render as Filters
     */
    items: JSX.Element[];
    filters: IFilter[];
}