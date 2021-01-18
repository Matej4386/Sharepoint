import * as React from 'react';
import * as strings from 'MspTableStrings';
import styles from '../FilterPanel.module.scss';
import {IMCollapseProps} from './IMCollapseProps';
import {IMCollapseState} from './IMCollapseState';

import {Link} from 'office-ui-fabric-react/lib/Link';
import {Label} from 'office-ui-fabric-react/lib/Label';
import {SearchBox} from 'office-ui-fabric-react/lib/SearchBox';
import {Icon} from 'office-ui-fabric-react/lib/Icon';
import { FilterTemplateOption } from '../../datatypes/FilterTemplateOption';

export default class MCollapse extends React.Component < IMCollapseProps, IMCollapseState > {
    constructor(props: IMCollapseProps) {
        super(props);
        this.state = {
            isCollapsed: props.filterField === 'all' ? props.filterConfiguration.showExpanded : true,
            valueFilter: '',
            showAll: false
        };
    }
    public componentDidUpdate (prevProps: IMCollapseProps): void {
        if (prevProps.filterField !== this.props.filterField) {
            this.setState({
                isCollapsed: this.props.filterField === 'all' ? this.props.filterConfiguration.showExpanded : true
            });
        }
    }
    public render(): React.ReactElement<IMCollapseProps> {
        let filteredItems: JSX.Element[] = [];
        if (this.state.valueFilter !== '') {
            filteredItems = [...this.props.items].filter(x => { return !this.isFilterMatch(x); });
        } else {
            filteredItems = [...this.props.items];
        }
        // if there is more than 15 items to show - slice them and show Shor more text
        if ((filteredItems.length > 15) && (this.state.showAll === false)) {
            filteredItems = filteredItems.slice(0, 15);
            const showAllElement: JSX.Element =
                <div style={{marginBottom: '0.5rem', marginTop: '0.5rem', paddingLeft: '27px'}}>
                    <Link
                        onClick={() => { this.setState({showAll: true}); }}
                    >
                        {strings.Filters.ShowAll}
                    </Link>
                </div>;
            filteredItems.push(showAllElement);
        }
        return (
            <div>
                <div className={styles.FilterPanelLayout__filterPanel__body__group__header}
                    style={this.props.groupIndex > 0 ? { marginTop: '10px' } : undefined}
                    onClick={() => {
                        const col: boolean = !this.state.isCollapsed;
                        this.setState({
                            isCollapsed: col
                        });
                    }}
                >
                    <div className={styles.FilterPanelLayout__filterPanel__body__headerIcon}>
                    {this.state.isCollapsed ?
                        <Icon iconName='ChevronDown' />
                        :
                        <Icon iconName='ChevronUp' />
                    }
                    </div>
                    <Label className='ms-font-l' style={{fontWeight: 'bold'}}>{this.props.filter.title.indexOf(' ') > -1 ? this.props.filter.title.split(' ')[2] : this.props.filter.title}</Label>
                </div>
                {((this.state.isCollapsed) &&
                <div>
                        {(this.props.filterConfiguration.showValueFilter) &&
                            <div
                                style={{
                                    marginTop: '0.7rem',
                                    marginBottom: '0.7rem',
                                    maxWidth: '10rem',
                                    paddingLeft: '27px'
                                }}
                            >
                                <SearchBox
                                    value={this.state.valueFilter}
                                    placeholder={strings.Filters.FilterPlacehoder}
                                    underlined={true}
                                    onChanged={(newValue?: string) => { this.onValueFilterChanged(newValue); }}
                                    onSearch={(newValue?: string) => { this.onValueFilterChanged(newValue); }}
                                    onClear={() => { this.setState({valueFilter: ''}); }}
                                    onClick={this.onValueFilterClick}
                                />
                            </div>
                        }
                    <div>
                        {
                            (this.props.filterConfiguration.filterMode === FilterTemplateOption.CheckBoxMulti) &&
                            (filteredItems.length > 5) &&
                            <div style={{marginBottom: '0.5rem', marginTop: '0.5rem', paddingLeft: '27px'}}>
                                <Link
                                    onClick={() => { this.props.onFilter(this.props.filter, true); }}
                                >
                                    {strings.Filters.ApplyFiltersLabel}
                                </Link>&nbsp;|&nbsp;<Link
                                    onClick={ () => {this.props.onFilter(this.props.filter, false); }}
                                >
                                    {strings.Filters.ClearFiltersLabel}
                                </Link>
                            </div>
                        }
                        {
                            filteredItems
                        }
                        {
                            (this.props.filterConfiguration.filterMode === FilterTemplateOption.CheckBoxMulti) &&
                            <div style={{marginTop: '0.7rem', marginBottom: '0.7rem', paddingLeft: '27px'}}>
                                <Link
                                    onClick={() => { this.props.onFilter(this.props.filter, true); }}
                                >
                                    {strings.Filters.ApplyFiltersLabel}
                                </Link>&nbsp;|&nbsp;<Link
                                    onClick={ () => {this.props.onFilter(this.props.filter, false); }}
                                >
                                    {strings.Filters.ClearFiltersLabel}
                                </Link>
                            </div>
                        }
                    </div>
                </div>
                )}
            </div>
        );
    }
    private isFilterMatch = (item: JSX.Element): boolean => {
        if (!this.state.valueFilter) {
            return false;
        }
        const isSelected: boolean = item.props.checked;
        if (isSelected) {
            return false;
        }
        return item.props.children.props.label.toLowerCase().indexOf(this.state.valueFilter.toLowerCase()) === -1;
    }
    private onValueFilterChanged = (newValue: string) => {
        this.setState({
            valueFilter: newValue
        });
    }
    private onValueFilterClick = (event: React.MouseEvent<HTMLInputElement | HTMLTextAreaElement>) => {
        event.stopPropagation();
    }
}