import {IFilterConfiguration} from '../../datatypes/IFilterConfiguration';
import {IFilter} from '../../datatypes/IFilter';
export interface IMCollapseProps {
    filterConfiguration: IFilterConfiguration;
    filter: IFilter;
    items: JSX.Element[];
    groupIndex: number;
    filterField: string;
    onFilter: (filterOption: IFilter, applyFilter: boolean) => void;
}