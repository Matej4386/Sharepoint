import * as React from 'react';
import styles from './FilterPanel.module.scss';
import * as strings from 'MspTableStrings';
import {IFilterPanelProps} from './IFilterPanelProps';
import {IFilterPanelState} from './IFilterPanelState';

import {FilterTemplateOption} from '../datatypes/FilterTemplateOption';
import {IFilters, IFilterVal} from '../datatypes/IFilters';
import MCollapse from './modules/MCollapse/MCollapse';
import MDatePicker from './modules/MDatePicker/MDatePicker';

import * as moment from 'moment';
import {isEqual} from '@microsoft/sp-lodash-subset';

import {ScrollablePane} from 'office-ui-fabric-react/lib/ScrollablePane';
import {Link} from 'office-ui-fabric-react/lib/Link';
import {Label} from 'office-ui-fabric-react/lib/Label';
import {Panel, PanelType} from 'office-ui-fabric-react/lib/Panel';
import {Checkbox} from 'office-ui-fabric-react/lib/Checkbox';

export default class FilterPanel extends React.Component < IFilterPanelProps, IFilterPanelState > {
    constructor(props: IFilterPanelProps) {
        super(props);
        this.state = {
            items: []
        };
    }
    public componentDidMount(): void {
        this.initItems(this.props);
    }
    public componentDidUpdate(prevProps: IFilterPanelProps, prevState: IFilterPanelState): void {
        if (this.props.debug === true) {
            console.log ('-------FilterPanel component-------');
            console.log ('Event: componentDidUpdate');
            console.log (this.props);
            console.log (this.state);
            console.log ('-----------------------------------');
        }
        if (!isEqual(prevProps, this.props)) {
            this.initItems(this.props);
        }
    }
    public render(): React.ReactElement<IFilterPanelProps> {
        const renderSelectedFilterValues: JSX.Element[] = this.props.selectedFilters.map((value: IFilters) => {
            let filtername: string = `[${value.filterTitle}: "`;
            if (value.filterValues.length > 0) {
                value.filterValues.map((filterval: IFilterVal) => {
                    if (value.filterType === 'DateTime') {
                        if (filterval.title.indexOf('|') > -1) {
                            const split: string[] = filterval.title.split('|');
                            if (split[0] !== 'null') {
                                filtername += moment(split[0]).format('L') + ' | ';
                            } else {
                                filtername += 'N/A | ';
                            }
                            if (split[1] !== 'null') {
                                filtername += moment(split[1]).format('L') + ' ';
                            } else {
                                filtername += 'N/A';
                            }
                        }
                    } else {
                        filtername += filterval.title + ' ';
                    }
                });
            }
            filtername += '"]';
            return (
              <Label className={styles.filter}>
                {filtername}
              </Label>
            );
        });
        const renderLinkRemoveAll: JSX.Element =
            (this.props.selectedFilters.length > 0) &&
                (<div className={`${styles.FilterPanelLayout__filterPanel__body__removeAllFilters}`}>
                    <Link onClick={this.removeAllFilters}>
                        {strings.Filters.RemoveAllFiltersLabel}
                    </Link>
                </div>);
        return (
            <Panel
                isOpen={ this.props.showPanel }
                type={ PanelType.medium }
                onDismiss={ this.onClosePanel }
                isLightDismiss={true}
                headerText={strings.Filters.FilterPanelTitle}
                closeButtonAriaLabel={strings.Filters.FilterPanelClose}
                onRenderFooterContent={ this.onRenderFooterContent }
                onRenderBody={() => {
                    if (this.props.filters.length > 0 ) {
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
                                        {renderLinkRemoveAll}
                                        {(this.props.selectedFilters.length > 0) &&
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
        props.filters.map((filter: IFilters) => {
            if ((props.filterField === 'all') || (props.filterField === filter.filterName)) {
                groupIndex++;
                let isSelected: boolean = false;
                let elements: JSX.Element[] = [];
                if (filter.filterMode === FilterTemplateOption.FixedDateRange) {
                    const isSelectedFilter: IFilters[] = this.props.selectedFilters.filter((selFilter) => selFilter.filterName === filter.filterName);
                    if (isSelectedFilter.length > 0) {
                        isSelected = true;
                    }
                    elements.push(
                        <MDatePicker
                            currentCultureName={this.props.currentCultureName}
                            filter={filter}
                            isSelectedItem={isSelected}
                            onFilter={this.props.onFilter}
                        />
                    );
                } else {
                    elements = filter.filterValues.map((filterValue: IFilterVal) => {
                        const isChecked: boolean = this.isValueInFilterSelection(filter, filterValue.title);
                        if (isChecked === true) {
                            isSelected = true;
                        }
                        let element: JSX.Element = undefined;
                        switch (filter.filterMode) {
                            case FilterTemplateOption.CheckBoxMulti:
                            case FilterTemplateOption.CheckBox:
                                element =
                                    <div className={styles.FilterPanelLayout__filterPanel__body__group__item}>
                                        <Checkbox
                                            label={filterValue.title + ' (' + filterValue.count + ')'}
                                            checked={isChecked}
                                            onChange={(ev, checked: boolean) => {
                                                filter.filterMode === FilterTemplateOption.CheckBoxMulti ?
                                                    checked ?
                                                        this.props.onFilter(filter, [filterValue.title], true, true, false)
                                                        :
                                                        this.props.onFilter(filter, [filterValue.title], false, true, false)
                                                    :
                                                    checked ?
                                                        this.props.onFilter(filter, [filterValue.title], true, false, false)
                                                        :
                                                        this.props.onFilter(filter, [filterValue.title], false, false, false);
                                            }}
                                        />
                                    </div>;
                                break;
                        }
                        return element;
                    });
                }
                items.push(
                    <div>
                            <MCollapse
                                filterField={props.filterField}
                                filter={filter}
                                items={elements}
                                groupIndex={groupIndex}
                                onFilter={this.props.onFilter}
                                isSelectedItem={isSelected}
                            />
                    </div>
                );
            }
        });
        this.setState({
            ...this.state,
            items: items
        });
    }
    private isValueInFilterSelection = (filter: IFilters, filterValue: string): boolean => {
        let isSelected: boolean = false;
        let found: IFilterVal[] = [];

        const isSelectedFilter: IFilters[] = this.props.selectedFilters.filter((selFilter) => {
            return selFilter.filterName === filter.filterName;
        });
        if (isSelectedFilter.length > 0) {
            found = isSelectedFilter[0].filterValues.filter((filValue) => {
                return filValue.title === filterValue;
            });
        }
        if (found.length > 0 ) {
            isSelected = true;
        }
        return isSelected;
    }
    private removeAllFilters = (): void => {
        this.props.onRemoveAllFilters();
    }
    private onClosePanel = (): void => {
        this.props.onUpdateShow(false);
    }
}