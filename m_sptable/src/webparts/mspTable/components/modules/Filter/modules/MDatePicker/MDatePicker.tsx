import * as React from 'react';
import * as strings from 'MspTableStrings';
import styles from './MDatePicker.module.scss';
import {IMDatePickerProps} from './IMDatePickerProps';
import {IMDatePickerState} from './IMDatePickerState';

import * as moment from 'moment';
import {isEqual } from '@microsoft/sp-lodash-subset';

import {DatePicker, IDatePickerProps} from 'office-ui-fabric-react/lib/DatePicker';
import {Link} from 'office-ui-fabric-react/lib/Link';

export default class MDatePicker extends React.Component < IMDatePickerProps, IMDatePickerState > {
    constructor(props: IMDatePickerProps) {
        super(props);
        this.state = {
            selectedFromDate: undefined,
            selectedToDate: undefined
        };
    }
    public componentDidMount (): void {
        if ((this.props.filter.filterValues.length > 0) && (this.props.filter.filterValues[0].title !== 'null')) {
            const splitRange: string[] = this.props.filter.filterValues[0].title.split('|');
            if (splitRange[0] !== 'null') {
                this.setState({
                    selectedFromDate: new Date(splitRange[0])
                });
            }
            if (splitRange[1] !== 'null') {
                this.setState({
                    selectedToDate: new Date(splitRange[1])
                });
            }
        }
    }
    public componentDidUpdate (prevProps: IMDatePickerProps): void {
        if (!isEqual(this.props.filter.filterValues, prevProps.filter.filterValues)) {
            if ((this.props.filter.filterValues.length > 0) && (this.props.filter.filterValues[0].title !== 'null')) {
                const splitRange: string[] = this.props.filter.filterValues[0].title.split('|');
                if (splitRange[0] !== 'null') {
                    this.setState({
                        selectedFromDate: new Date(splitRange[0])
                    });
                }
                if (splitRange[1] !== 'null') {
                    this.setState({
                        selectedToDate: new Date(splitRange[1])
                    });
                }
            } else if (this.props.filter.filterValues[0].title === 'null') {
                this.setState({
                    selectedFromDate: undefined,
                    selectedToDate: undefined
                });
            }
        }
    }
    public render(): React.ReactElement<IMDatePickerProps> {
        const fromProps: IDatePickerProps = {
            placeholder: strings.Templates.DateFromLabel,
            onSelectDate: this.updateFromDate,
            value: this.state.selectedFromDate,
            formatDate: this.onFormatDate,
            showGoToToday: true,
            borderless: true,
            strings: strings.Templates.DatePickerStrings
        };

        const toProps: IDatePickerProps = {
            placeholder: strings.Templates.DateTolabel,
            onSelectDate: this.updateToDate,
            value: this.state.selectedToDate,
            formatDate: this.onFormatDate,
            showGoToToday: true,
            borderless: true,
            strings: strings.Templates.DatePickerStrings
        };
        if (this.state.selectedFromDate) {
            const minDdate: Date = new Date(this.state.selectedFromDate.getTime());
            minDdate.setDate(this.state.selectedFromDate.getDate() + 1);
            toProps.minDate = minDdate;
        }

        if (this.state.selectedToDate) {
            const maxDate: Date = new Date(this.state.selectedToDate.getTime());
            maxDate.setDate(this.state.selectedToDate.getDate() - 1);
            fromProps.maxDate = maxDate;
        }
        return (
            <div style={{paddingLeft: '15px', maxWidth: '15rem'}} className={styles.MDatePicker}>
                <DatePicker {...fromProps} />
                <DatePicker {...toProps} />
                <div style={{marginTop: '0.7rem', marginBottom: '0.7rem', paddingLeft: '15px'}}>
                    <Link
                        onClick={() => {
                            this.setState({
                                selectedFromDate: undefined,
                                selectedToDate: undefined
                            });
                            this.props.onFilter(
                                this.props.filter,
                                [this.props.filter.filterValues[0].title],
                                false,
                                false,
                                false
                            );
                        }}
                        disabled={!(this.state.selectedToDate || this.state.selectedFromDate)}
                    >
                        {strings.Filters.ClearFiltersLabel}
                    </Link>
                </div>
            </div>
        );
    }
    private updateFromDate = (fromDate: Date) => {
        this.setState({
            selectedFromDate: fromDate
        });
        this.updateFilter(fromDate, this.state.selectedToDate);
    }
    private updateToDate = (toDate: Date) => {
        this.setState({
            selectedToDate: toDate
        });
        this.updateFilter(this.state.selectedFromDate, toDate);
    }
    private updateFilter = (selectedFromDate: Date, selectedToDate: Date) => {
        const startDate: string = selectedFromDate ? selectedFromDate.toISOString() : 'null';
        const endDate: string = selectedToDate ? selectedToDate.toISOString() : 'null';
        const rangeConditions: string = `${startDate}|${endDate}`;

        this.props.onFilter(this.props.filter, [rangeConditions], true, false, false);
    }
    private onFormatDate = (date: Date): string => {
        moment.locale(this.props.currentCultureName);
        return moment(date).format('LL');
    }
}