import * as React from 'react';
import styles from './FilterPanel.module.scss';
import * as strings from 'MspTableStrings';
import {IFilterPanelProps} from './IFilterPanelProps';
import {IFilterPanelState} from './IFilterPanelState';

import {IFilterOptions} from '../datatypes/IFilterOptions';
import {IFilterConfiguration} from '../datatypes/IFilterConfiguration';
import {IFilter} from '../datatypes/IFilter';
import {FilterTemplateOption} from '../datatypes/FilterTemplateOption';
import MCollapse from './MCollapse/MCollapse';
import {isEqual} from '@microsoft/sp-lodash-subset';

import * as $ from 'jquery';
import {ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {ScrollablePane} from 'office-ui-fabric-react/lib/ScrollablePane';
import {IconButton, ActionButton} from 'office-ui-fabric-react/lib/Button';
import {Label} from 'office-ui-fabric-react/lib/Label';
import {Panel, PanelType} from 'office-ui-fabric-react/lib/Panel';
import {Checkbox} from 'office-ui-fabric-react/lib/Checkbox';


export default class FilterPanel extends React.Component < IFilterPanelProps, IFilterPanelState > {
    constructor(props: IFilterPanelProps) {
        super(props);
        this.state = {
            items: [],
            filters: []
        };
    }
    public async componentWillMount (): Promise<any> {
        this.loadFilters();
    }
    public async componentDidUpdate(prevProps: IFilterPanelProps, prevState: IFilterPanelState): Promise<any> {
        /*
        if (this.props.debug === true) {
            console.log ('-------FilterPanel component-------');
            console.log ('Event: componentDidUpdate');
            console.log (this.props);
            console.log (this.state);
            console.log ('-----------------------------------');
        }*/
        if (!isEqual(prevProps.filtersConfiguration, this.props.filtersConfiguration)) {
            this.loadFilters();
        }
        if (!isEqual(prevProps.filterField,this.props.filterField)) {
            if (this.props.filterField === 'reload') {
                this.loadFilters();
            } else {
                this.initItems(this.props);
            }
        }
    }
    public render(): React.ReactElement<IFilterPanelProps> {
        const renderSelectedFilterValues: JSX.Element[] = this.state.filters.map((value: IFilter) => {
            let filtername: string = `[${value.title.split(' ')[2]}: "`;
            const selectedFiltersOptions = value.options.filter((filterval) => filterval.executed === true);
            selectedFiltersOptions.map((filterval) => {
                filtername += filterval.text + ' ';
            });
            filtername += '"]';
            if (selectedFiltersOptions.length > 0) {
                return (
                    <ActionButton
                        iconProps={ { iconName: 'ClearFilter' } }
                        ariaLabel={strings.Filters.ClearFiltersLabel}
                        onClick={() => this.onFilter(value, false)}
                    >
                        {filtername}
                    </ActionButton>
                  );
            } else {
                return null;
            }

        });
        return (

            <Panel
                isOpen={ this.props.showPanel }
                type={ PanelType.medium }
                onDismiss={ this.onClosePanel }
                isLightDismiss={true}
                hasCloseButton={false}
                headerText={strings.Filters.FilterPanelTitle}
                closeButtonAriaLabel={strings.Filters.FilterPanelClose}
                onRenderHeader={ this.onRenderHeader}
                onRenderFooterContent={ this.onRenderFooterContent }
                onRenderBody={() => {
                    if (this.props.filtersConfiguration.length > 0 ) {
                        return (
                            <div
                                style={{
                                    height: '100%',
                                    position: 'relative',
                                    maxHeight: 'inherit'
                                }}
                            >
                                <ScrollablePane>
                                    <div
                                        className={styles.FilterPanelLayout__filterPanel__body}
                                        data-is-scrollable={true}
                                    >
                                        {(renderSelectedFilterValues.length > 0) &&
                                            <div className={styles.FilterPanelLayout__selectedFilters}>
                                                {renderSelectedFilterValues}
                                            </div>
                                        }
                                        {this.state.items}
                                    </div>
                                </ScrollablePane>
                            </div>
                        );
                    }
                }}
            />
        );
    }
    private onRenderHeader = (): JSX.Element => {
        const isSelected: boolean = this.state.filters.filter((filter: IFilter) => {
            return filter.options.filter((option) => option.executed === true).length > 0;
        }).length > 0;
        return (
            <div className={styles.FilterPanelHeader}>
                <Label className={styles.FilterPanelHeaderText}>{strings.Filters.FilterPanelTitle}</Label>
                <IconButton
                    iconProps={ { iconName: 'ClearFilter' } }
                    title={strings.Filters.RemoveAllFiltersLabel}
                    ariaLabel={strings.Filters.RemoveAllFiltersLabel}
                    disabled={!isSelected}
                    onClick={() => this.onFilter(undefined, false)}
                />
                <IconButton
                    iconProps={ { iconName: 'ChromeClose' } }
                    title={strings.Filters.FilterPanelClose}
                    ariaLabel={strings.Filters.FilterPanelClose}
                    onClick={this.onClosePanel}
                />
            </div>
          );
    }
    private onRenderFooterContent = (): JSX.Element => {
        /*return (
          <div>
            <PrimaryButton
              onClick={ this._onClosePanel }
              style={{ 'marginRight': '8px' }}
            >
              {strings.ContractsView.Filter.FilterPanelClose}
            </PrimaryButton>
          </div>
        );*/
        return;
    }
    private initItems = (props: IFilterPanelProps): void => {
        const items: JSX.Element[] = [];
        let groupIndex: number = 0;
        if (props.filtersConfiguration) {
            props.filtersConfiguration.map((filterConfig: IFilterConfiguration) => {
                if ((props.filterField === 'all') || (props.filterField === filterConfig.filterName)) {
                    groupIndex++;
                    let elements: JSX.Element[] = [];
                    const filter: IFilter = this.state.filters.filter((fOptions: IFilter) => {
                        return fOptions.fieldName === filterConfig.filterName;
                    })[0];
                    // we have loaded filterOption from api
                    if (filter) {
                        elements = filter.options.map((options) => {
                            let element: JSX.Element = undefined;
                            element =
                                <div className={styles.FilterPanelLayout__filterPanel__body__group__item}>
                                    <Checkbox
                                        label={options.text}
                                        value={options.value}
                                        checked={options.selected}
                                        onChange={(ev, checked: boolean) => {
                                            if (filterConfig.filterMode === FilterTemplateOption.CheckBoxMulti) {
                                                const stateFilters: IFilter[] = [...this.state.filters];
                                                const obj: IFilter = stateFilters.find(stateFilter => stateFilter.fieldName === filter.fieldName);
                                                if (obj) {
                                                    const obj2 = obj.options.find(option => option.value === options.value);
                                                    if (obj2) {
                                                        obj2.selected = checked;
                                                    }
                                                }
                                                this.setState({
                                                    filters: stateFilters
                                                }, () => this.initItems(this.props));
                                                // Checkbox not multi, filtering is fired as soon as option is checked
                                            } else {
                                                const stateFilters: IFilter[] = [...this.state.filters];
                                                const obj: IFilter = stateFilters.find(stateFilter => stateFilter.fieldName === filter.fieldName);
                                                if (obj) {
                                                    const obj2 = obj.options.find(option => option.value === options.value);
                                                    if (obj2) {
                                                        obj2.selected = checked;
                                                        obj2.executed = checked;
                                                    }
                                                }
                                                this.setState({
                                                    filters: stateFilters
                                                }, () => {
                                                    this.props.onFilter(this.state.filters);
                                                    this.initItems(this.props);
                                                });
                                            }

                                        }}
                                    />
                                </div>;
                            return element;
                        });
                        items.push(
                            <div>
                                <MCollapse
                                    filterField={props.filterField}
                                    filterConfiguration={filterConfig}
                                    filter={filter}
                                    items={elements}
                                    groupIndex={groupIndex}
                                    onFilter={this.onFilter}
                                />
                            </div>
                        );
                    }
                }
            });
        }
        this.setState({
            ...this.state,
            items: items
        });
    }
    // for multi valued filters
    private onFilter = (filter: IFilter, applyFilter: boolean): void => {
        if (applyFilter === true) {
            const stateFilters: IFilter[] = [...this.state.filters];
            const obj: IFilter = stateFilters.find(stateFilter => stateFilter.fieldName === filter.fieldName);
            if (obj) {
                const obj2 = obj.options.map(option => ( option.selected === true ? {...option, executed: true} : option));
                obj.options = obj2;
            }
            this.setState({
                filters: stateFilters
            }, () => {
                this.props.onFilter(this.state.filters);
                this.initItems(this.props);
            });
        } else {
            // clear all
            if (filter === undefined) {
                const stateFilters: IFilter[] = this.state.filters.map((filter2) => {
                    const newFilter: IFilter = filter2;
                    const obj2 = newFilter.options.map(option => ( option.selected === true ? {...option, executed: false, selected: false} : option));
                    newFilter.options = obj2;
                    return newFilter;
                });
                this.setState({
                    filters: stateFilters
                }, () => {
                    this.props.onFilter(this.state.filters);
                    this.initItems(this.props);
                });
            //  unselect filter option
            } else {
                const stateFilters: IFilter[] = [...this.state.filters];
                const obj: IFilter = stateFilters.find(stateFilter => stateFilter.fieldName === filter.fieldName);
                if (obj) {
                    const obj2 = obj.options.map(option => ( option.selected === true ? {...option, executed: false, selected: false} : option));
                    obj.options = obj2;
                }
                this.setState({
                    filters: stateFilters
                }, () => {
                    this.props.onFilter(this.state.filters);
                    this.initItems(this.props);
                });
            }
        }
    }
    private onClosePanel = (): void => {
        this.props.onUpdateShow(false);
    }
    private loadFilters = async (): Promise<any> => {
        if (this.props.filtersConfiguration) {
            const filters: IFilter[] = [];
            for (let i: number = 0; i < this.props.filtersConfiguration.length; i++) {
                const data: IFilterOptions = await this.getFiltersForField(this.props.filtersConfiguration[i]);
                // if undefined - unsupoorted filtering
                if (data.fieldName !== undefined) {
                    const newData: IFilter = {
                        title: data.title,
                        fieldName: data.fieldName,
                        options: []
                    };
                    for (let j: number = 0; j < data.options.length; j++) {
                        newData.options.push({
                            value: data.options[j].value,
                            text: data.options[j].text,
                            selected: data.options[j].selected,
                            executed: false
                        });
                    }
                    filters.push(newData);
                }
            }
            this.setState({
                filters: [...filters]
            }, () => this.initItems(this.props));
        }
    }
    private getFiltersForField = async (filterConfiguration: IFilterConfiguration): Promise<any> => {
        return new Promise<any>((resolve, reject) => {
            const httpClientOptions: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-type': 'application/json;odata=verbose',
                    'X-SP-REQUESTRESOURCES': 'listUrl=' + encodeURIComponent(this.props.listUrl),
                    'odata-version': ''
                }
            };
            const endpoint: string = `${this.props.webAbsoluteUrl}/_api/web/GetList(@listUrl)/RenderListFilterData`
                + `?@listUrl=${encodeURIComponent('\'' + this.props.listUrl + '\'')}`
                + `&FieldInternalName=${filterConfiguration.filterName}`
                + `&ViewId=${this.props.currentViewId}`;
            this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, httpClientOptions)
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    return response.text(); // text not json
                } else {
                    reject('Error: get filters');
                }
            })
                .then((data) => {
                    /**
                     * Standard response contains html select with options
                     * Parsin via regexp can be done also by solution below via DOMParser
                     */
                    if (data.indexOf('<OPTION') !== -1) {
                        let regexp: RegExp = /(<OPTION[^>]*?>([^<]+)<\/OPTION>)/g;
                        const getMatches = (string: string, regex: RegExp, index: number): any[] => {
                            index || (index = 1); // default to the first capturing group
                            const mymatches: any[] = [];
                            let match: any;
                            while (match = regex.exec(string)) {
                            mymatches.push(match[index]);
                            }
                            return mymatches;
                        };
                        const matches: any[] = getMatches(data, regexp, 1);
                        const objects: any[] = [];
                        for (let i: number = 0; i < matches.length; i++) {
                            const obj: any = {
                                value: '',
                                text: '',
                                selected: false
                            };
                            regexp = /(?<=\")(.*?)(?=\")/g;
                            obj.value = regexp.exec(matches[i])[0];
                            regexp = /(?<=\>)(.*?)(?=\<)/g;
                            obj.text = regexp.exec(matches[i])[0];
                            obj.selected = matches[i].indexOf('SELECTED') !== -1 ? true : false;
                            if (obj.text !== '(Všetko)') {
                                if (obj.text !== '(All)') {
                                    objects.push(obj);
                                }
                            }
                        }
                        regexp = /(?<=\TITLE=")(.*?)(?=\")/g;
                        const title: string = regexp.exec(data)[0];
                        regexp = /(?<=}",")(.*?)(?=\")/g;
                        const fieldName: string = regexp.exec(data)[0];
                        const finalObject: IFilterOptions = {
                            title: title, // title should contains 3 strings split by ' ', in rendering we are spliting title by ' ' and taking 3. string
                            fieldName: fieldName,
                            options: objects
                        };
                        resolve(finalObject);
                        /**
                         * Filters for Managed metadata are returned as UL LI html
                         * Parsing via DOMParser
                         * Response does not contains DisplayName for Filter so we are passing Columns definition to extract Display name
                         */
                    } else if (data.indexOf('<ul') !== -1) {
                        const parser: DOMParser = new DOMParser();
                        const parsedHtml: Document = parser.parseFromString(data, 'text/html');
                        const mylist = [];
                        let counter: number = 0;
                        $(parsedHtml).find('li').each(function () {
                            counter++;
                            if (counter === 2) {
                                mylist.push({
                                    value: $(this).text(),
                                    text: $(this).text(),
                                    selected: false
                                });
                            }
                            if (counter === 5) {
                                counter = 0;
                            }
                        });
                        const columnDef: any = this.props.columnsDefinition.filter((columnDefTemp) => filterConfiguration.filterName === columnDefTemp.RealFieldName)[0];
                        if (columnDef !== undefined) {
                            const finalObject: IFilterOptions = {
                                title: 'Filter for ' + columnDef.DisplayName, // title should contains 3 strings split by ' ', in rendering we are spliting title by ' ' and taking 3. string
                                fieldName: filterConfiguration.filterName,
                                options: mylist
                            };
                            resolve(finalObject);
                        }

                    } else {
                        const finalObject: IFilterOptions = {
                            title: filterConfiguration.filterName + '-Unsupported',
                            fieldName: filterConfiguration.filterName,
                            options: []
                        };
                        resolve(finalObject);
                    }
                })
            .catch((error) => {
                reject('Error: get filters: ' + error);
            });
        });
    }
}