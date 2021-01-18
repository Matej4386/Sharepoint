import {IFilterConfiguration} from '../datatypes/IFilterConfiguration';
import {IFilterOptions} from '../datatypes/IFilterOptions';
import {SPHttpClient} from '@microsoft/sp-http';
export interface IFilterPanelProps {
    debug: boolean;
    webAbsoluteUrl: string;
    currentViewId: string;
    listUrl: string;
    spHttpClient: SPHttpClient;
    showPanel: boolean;
    filtersConfiguration: IFilterConfiguration[];
    filterField: string;
    renderInfo: (error: boolean, message: string) => void;
    onUpdateShow: (show: boolean) => void;
    onFilter: (filtersOption: IFilterOptions[]) => void;
    // only for managed metadata filters - resp api does not return field display name
    columnsDefinition: any[];
}