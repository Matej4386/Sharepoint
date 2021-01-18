import * as React from 'react';
import styles from './MspTable.module.scss';
import * as strings from 'MspTableStrings';
import {IMspTableProps} from './IMspTableProps';
import {IMspTableState} from './IMspTableState';

import {IFilters, IFilterVal} from './modules/datatypes/IFilters';
import {FilterTemplateOption} from './modules/datatypes/FilterTemplateOption';
import {IFilterConfiguration} from './modules/datatypes/IFilterConfiguration';
import {FiltersSortOption} from './modules/datatypes/FiltersSortOptions';
import {FiltersSortDirection} from './modules/datatypes/FiltersSortDirection';

import  FilterPanel  from './modules/Filter/FilterPanel';

import * as $ from 'jquery';
import * as moment from 'moment';
import {sp} from '@pnp/sp';
import '@pnp/polyfill-ie11';
/**
 * Creates an array of array values not included in the other given arrays using SameValueZero for equality comparisons. 
 * The order and references of result values are determined by the first array.
 */
import {difference} from '@microsoft/sp-lodash-subset';
/**
 * A utility for converting HTML strings into React components. 
 * Avoids the use of dangerouslySetInnerHTML and converts standard HTML elements, attributes and inline styles into their React equivalents.
 */
import ReactHtmlParser from 'react-html-parser';

import {CommandBar} from 'office-ui-fabric-react/lib/CommandBar';
import {Sticky, StickyPositionType} from 'office-ui-fabric-react/lib/Sticky';
import {SearchBox} from 'office-ui-fabric-react/lib/SearchBox';
import {IRenderFunction} from 'office-ui-fabric-react/lib/Utilities';
import {ScrollablePane} from 'office-ui-fabric-react/lib/ScrollablePane';
import {Persona, PersonaSize} from 'office-ui-fabric-react/lib/Persona';
import {IContextualMenuProps, IContextualMenuItem, DirectionalHint, ContextualMenu} from 'office-ui-fabric-react/lib/ContextualMenu';
import {CheckboxVisibility, IDetailsHeaderProps, ColumnActionsMode, ConstrainMode, DetailsListLayoutMode, IColumn, Selection, buildColumns} from 'office-ui-fabric-react/lib/DetailsList';
import {ShimmeredDetailsList} from 'office-ui-fabric-react/lib/ShimmeredDetailsList';

export default class MspTable extends React.Component<IMspTableProps, IMspTableState> {
  private gridItems: any = undefined;
  private selection: Selection;
  constructor(props: IMspTableProps) {
    super(props);
    moment.locale(props.currentCultureName);

    this.selection = new Selection({
        onSelectionChanged: this.onItemsSelectionChanged
    });
    this.selection.setItems([], false);
    this.state = {
        items: [],
        columnsDefinition: [],
        columns: this.buildColumns(
            [], 
            [],
            true,
            this.onColumnClick,
            '',
            undefined,
            this.onColumnContextMenu
        ),
        filters: [],
        selectedFilters: [],
        showFiltersPanel: false,
        filterField: '',
        contextualMenuProps: undefined,
        sortedColumnKey: props.onInitSortedKey,
        isSortedDescending: false,
        isLoadingData: false,
        valueFilter: '',
        editValueText: '',
        selectionDetails: [],
        commandBarItems: [],
        farItems: []
    };
    // https://tc39.github.io/ecma262/#sec-array.prototype.findindex IE11 polyfill
    if (!Array.prototype.findIndex) {
      Object.defineProperty(Array.prototype, 'findIndex', {
        value: function(predicate) {
      // 1. Let O be ? ToObject(this value).
          if (this == null) {
              throw new TypeError('"this" is null or not defined');
          }

          var o = Object(this);

          // 2. Let len be ? ToLength(? Get(O, "length")).
          var len = o.length >>> 0;

          // 3. If IsCallable(predicate) is false, throw a TypeError exception.
          if (typeof predicate !== 'function') {
              throw new TypeError('predicate must be a function');
          }

          // 4. If thisArg was supplied, let T be thisArg; else let T be undefined.
          var thisArg = arguments[1];

          // 5. Let k be 0.
          var k = 0;

          // 6. Repeat, while k < len
          while (k < len) {
          // a. Let Pk be ! ToString(k).
          // b. Let kValue be ? Get(O, Pk).
          // c. Let testResult be ToBoolean(? Call(predicate, T, « kValue, k, O »)).
          // d. If testResult is true, return k.
          var kValue = o[k];
          if (predicate.call(thisArg, kValue, k, o)) {
              return k;
          }
          // e. Increase k by 1.
          k++;
          }

          // 7. Return -1.
          return -1;
        },
        configurable: true,
        writable: true
      });
    }
  }
  public componentDidMount(): void {
    this.readDataForGrid(undefined);
  }
  public componentDidUpdate(prevProps: IMspTableProps, prevState: IMspTableState): void {
    if (this.props.debug === true) {
        console.log ('-------MspTable component-------');
        console.log ('Event: componentDidUpdate');
        console.log (this.props);
        console.log (this.state);
        console.log ('-----------------------------------');
    }
    if (prevProps.listToDisplayId !== undefined) {
        if (prevProps.listToDisplayId !== this.props.listToDisplayId) {
          this.readDataForGrid(undefined);
        }
    }
    if ((prevProps.listToDisplayId === undefined) && (this.props.listToDisplayId !== undefined)) {
      this.readDataForGrid(undefined);
    }
    if (prevProps.listToDisplayId !== this.props.listToDisplayId) {
      this.readDataForGrid(undefined);
    }
}
  public render(): React.ReactElement<IMspTableProps> {
    let renderGrid: JSX.Element;
    let renderCommandBar: JSX.Element;

    if (this. props.commandBar === true) {
      renderCommandBar =
      <div className={styles.CommandBarcontainerMaxWidth}>
          <CommandBar
            isSearchBoxVisible={false}
            items={this.state.commandBarItems}
            farItems={this.state.farItems}
          />
        </div>;
    }

    renderGrid =
    <div>
      <ShimmeredDetailsList
        setKey='items'
        items={this.state.items!}
        selection={this.selection}
        columns={this.state.columns}
        checkboxVisibility={CheckboxVisibility.onHover}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        isHeaderVisible={true}
        selectionMode={this.props.selectionMode }
        constrainMode={ConstrainMode.unconstrained}
        enterModalSelectionOnTouch={true}
        onItemInvoked={this.onItemInvoked}
        onItemContextMenu={this.onItemContextMenu}
        onColumnHeaderClick={this.onColumnClick}
        onRenderDetailsHeader={this.onRenderDetailsHeader}
        enableShimmer={this.state.isLoadingData}
        selectionZoneProps={{
            selection: this.selection,
            disableAutoSelectOnInputElements: true,
            selectionMode: this.props.selectionMode
        }}
      />
      <FilterPanel
        showPanel={this.state.showFiltersPanel}
        onUpdateShow={(show) => this.setState({showFiltersPanel: show})}
        debug={this.props.debug}
        renderInfo={this.props.renderInfo}
        filters={this.state.filters}
        selectedFilters={this.state.selectedFilters}
        onFilter={this.onFilter}
        filterField={this.state.filterField}
        onRemoveAllFilters={this.onRemoveAllFilters}
        currentCultureName={this.props.currentCultureName}
      />
    </div>;
    return(
      <div>
        {renderCommandBar}
        <div
          className={styles.viewcontainer}
          onScroll={this.handleScroll}
        >
          <ScrollablePane className={styles.scrollable}>
            <div className={styles.containerChild}></div>
              <div>
                <div>
                  {renderGrid}
                </div>
                {this.state.contextualMenuProps && <ContextualMenu {...this.state.contextualMenuProps} />}
              </div>
          </ScrollablePane>
        </div>
      </div>
    );
  }
  private handleScroll(event: any) {
    const element = document.querySelector("[class*='stickyAbove-']");
    if (element != null) {
        element.scrollLeft = event.target.scrollLeft;
    }
  }
  private onItemsSelectionChanged = (): void => {
    this.setState({ selectionDetails: this.getSelectionDetails() });
  }
  private onItemInvoked = (item: any, index: number): void => {
    /**
     * Place the code for row double click here or pass function to props
     */
    console.log ('Item invoked');
    console.log (item);
    console.log (index);
  }
  private buildColumns(
    items: any[],
    columnsDefinition: any[],
    canResizeColumns?: boolean,
    onColumnClick?: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => any,
    sortedColumnKey?: string,
    isSortedDescending?: boolean,
    onColumnContextMenu?: (column: IColumn, ev: React.MouseEvent<HTMLElement>) => any
  ): any {
    const columns = buildColumns(
        items,
        canResizeColumns,
        onColumnClick,
        sortedColumnKey,
        isSortedDescending
    );
    const gridColumns: IColumn[] = [];

    columns.forEach(column => {
      const columnDef = columnsDefinition.filter((columnDefTemp) => column.fieldName === columnDefTemp.Name)[0];
      if (columnDef !== undefined) {

          column.maxWidth = 200;
          column.name = columnDef.DisplayName;
          if (columnDef.Type === 'Note') {
            column.isMultiline = true;
            column.maxWidth = 200;
            column.onRender = (item: any) => (
                ReactHtmlParser(item[column.fieldName])
            );
          }
          if (columnDef.Type === 'Lookup') { //Taxonomy Lookup - be aware of connected lists (to view connected list columns change code)
            column.onRender = (item: any) => (
                item[column.fieldName].Label
            );
          }
          if (columnDef.Type === 'User') {
            column.maxWidth = 180;
            column.onRender = (item: any) => (
                <Persona
                    text={item[column.fieldName] ? item[column.fieldName][0].title : 'N/A'}
                    size={ PersonaSize.size10 }
                />
            );
          }
          column.isCollapsable = false;
          column.onColumnContextMenu = onColumnContextMenu;
          column.columnActionsMode = ColumnActionsMode.hasDropdown;

        if (column) {
          gridColumns.push(column);
        }
      }
    });
    return gridColumns;
  }
  private async readDataForGrid(viewId: string): Promise<any> {
    try {
      let rawData: any = undefined;
        if ((this.props.listToDisplayId !== undefined)) {
            this.setState({ ...this.state, isLoadingData: true });
            if (viewId !== undefined) {
                const overriderenderListDataParams: any = {
                    View: viewId
                };
                const renderListDataParams: any = {
                    RenderOptions: 1687
                };
                rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream(renderListDataParams, overriderenderListDataParams);
            } else {
                const renderListDataParams: any = {
                    RenderOptions: 1687
                };
                rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream(renderListDataParams);
            }
        }
        if (this.props.debug === true) {
          console.log(rawData);
        }
        this.setState({
          ...this.state,
          selectedFilters: [],
          valueFilter: '',
          items: (rawData) ?
            this.copyAndSort(rawData.ListData.Row, this.state.sortedColumnKey, this.state.isSortedDescending)
            :
            [],
          columnsDefinition: (rawData) ? 
            rawData.ListSchema.Field 
            : 
            [],
          columns: (rawData) ?
            this.buildColumns(rawData.ListData.Row, rawData.ListSchema.Field, true, this.onColumnClick, this.state.sortedColumnKey, this.state.isSortedDescending)
            :
            []
        });
        this.gridItems = (rawData) ? 
          rawData.ListData.Row 
          : 
          [];
        const farCommandBarItems: IContextualMenuItem[] = [];
        farCommandBarItems.push({
          id: 'resultsCount',
          key: 'resultsCount',
          name: strings.Table.ResultsCount + ' : ' + this.gridItems.length,
          iconProps: { iconName: 'Info' }
        });
        farCommandBarItems.push({
          key: 'search',
          onRender: () => <SearchBox
            placeholder={strings.Table.SearchPlaceholder}
            value={this.state.valueFilter}
            onChange={(newValue?: string) => this.onValueFilterChanged2(newValue)}
            onSearch={(newValue?: string) => this.onValueFilterChanged(newValue)}
            onClear={() =>  this.onValueFilterCleared()}
            onClick={this.onValueFilterClick}
            underlined={false}
          />
        });
        if (this.props.listToDisplayId !== undefined) {
          const viewData = await sp.web.lists.getById(this.props.listToDisplayId).views.filter('Hidden ne true').get();
          if (this.props.debug === true) console.log (viewData);
          // if we have more than default view for the list render option to change the view
          if (viewData.length > 1) {
            const submenu: IContextualMenuItem[] = [];
            for (let i = 0; i < viewData.length; i++) {
              submenu.push({
                key: i.toString(),
                name: viewData[i].Title,
                canCheck: true,
                checked: rawData.ViewMetadata.Id === viewData[i].Id,
                onClick: () => {this.readDataForGrid(viewData[i].Id); }
              });
            }
            farCommandBarItems.push({
              key: 'views',
              name: strings.Table.ChangeView,
              iconProps: { iconName: 'List' },
              subMenuProps: {
                  items: submenu
              }
            });
          }
        }
        const filters: IFilters[] = (rawData) ? this.createFilters(rawData.ListData.Row, undefined) : [];
        if (filters.length > 0) {
          farCommandBarItems.push({
            key: 'filters',
            name: strings.Table.filterText,
            iconProps: { iconName: 'Filter' },
            onClick: () => this.setState({...this.state, showFiltersPanel: true, filterField: 'all'})
          });
          this.setState({
            ...this.state,
            filters: filters,
            isLoadingData: false,
            farItems: (this.props.commandBar === true) ? farCommandBarItems : []
          });
        } else {
          this.setState({
            ...this.state,
            isLoadingData: false,
            farItems: (this.props.commandBar === true) ? farCommandBarItems : []
          });
        }
        if (farCommandBarItems.length > 0) {
          if (this.props.onCommandBarChange) {
            this.props.onCommandBarChange(undefined, farCommandBarItems);
          }
        }
    } catch (error) {
      const errorText = `MspTable (readDataForGrid) -> ${strings.Errors.ErrorLoadingData}: ${error}`;
      this.setState({ ...this.state, items: {}, isLoadingData: false });
      this.props.renderInfo (true, errorText);
    }
  }
  private search = (items: any, searchValue: string): any => {
    let searches: any[] = [];
    let myitems: any[] = [...items];
    const searchesArray: string[] = (this.props.searchConfiguration.length > 0) ?
      (this.props.searchConfiguration.indexOf(',') > -1 ) ?
        this.props.searchConfiguration.split(',')
        :
        [this.props.searchConfiguration]
      :
      [];
    // check if search column is in columns definition for the view
    if (searchValue !== '') {
      for (let j: number = 0; j < searchesArray.length; j++) {
        let parcialSearch: any[] = [];
        const colDef: any[] = this.state.columnsDefinition.filter((columnsDef: any) => columnsDef.RealFieldName === searchesArray[j]);
        if (colDef.length > 0) {
          parcialSearch = myitems.filter((i: any) => {
            if ((i[searchesArray[j]] === null) || (i[searchesArray[j]] === '')) {
                return false;
            }
            if (i[searchesArray[j]].toLowerCase().indexOf(searchValue.toLowerCase()) > -1) {
                return true;
            }
            return false;
          });
          myitems = difference(myitems, parcialSearch);
        }
        searches = [...searches, ...parcialSearch]; // concat
      }
    } else {
      searches = myitems;
    }
    return searches;
  }
  private onFilter = (filter: IFilters, filterValue: string[], add: boolean, multi: boolean, executeSearch: boolean): void => {
    if (executeSearch === true) {
      const selectedFilters: IFilters[] = this.state.selectedFilters;
      selectedFilters[selectedFilters.findIndex(p => p.filterName === filter.filterName)].executeSearch = true;
      const newItems: any = this.filter(this.state.selectedFilters, this.state.valueFilter);
      this.setState({
        selectedFilters: selectedFilters,
        items: newItems,
        columns: this.columnsFilter (this.state.selectedFilters),
        filters: this.createFilters(newItems, undefined)
      });
    } else {
      const myfilterValues: IFilterVal[] = [];
      for (let i: number = 0; i < filterValue.length; i++) {
        myfilterValues.push({
            title: filterValue[i],
            count: 0
        });
      }
      if (add === true) {
        const selectedFilters: IFilters[] = this.state.selectedFilters;
        let found: boolean = false;
        for (let i: number = 0; i < selectedFilters.length; i++) {
          if (selectedFilters[i].filterName === filter.filterName) {
            found = true;
            if (filter.filterMode === FilterTemplateOption.FixedDateRange) {
              selectedFilters[i].filterValues = myfilterValues;
            } else {
              selectedFilters[i].filterValues.push(...myfilterValues);
            }
          }
        }
        if (found === false) {
          selectedFilters.push({
            filterName: filter.filterName,
            filterTitle: filter.filterTitle,
            filterType: filter.filterType,
            filterValues: myfilterValues,
            filterMode: filter.filterMode,
            showExpanded: filter.showExpanded,
            showValueFilter: filter.showValueFilter,
            executeSearch: false
          });
        }
        if (multi === true) {
            this.setState({
                selectedFilters: selectedFilters
            });
        } else {
          const newItems: any = this.filter(selectedFilters, this.state.valueFilter);
          this.setState({
            selectedFilters: selectedFilters,
            items: newItems,
            columns: this.columnsFilter(selectedFilters),
            filters: this.createFilters(newItems, selectedFilters)
          });
        }
      } else {
        let selectedFilters: IFilters[] = this.state.selectedFilters;
        let found: boolean = false;
        if (multi === true && myfilterValues.length > 0) {
          for (let i: number = 0; i < selectedFilters.length; i++) {
            if (selectedFilters[i].filterName === filter.filterName) {
              if (selectedFilters[i].filterValues.length > 1) {
                found = true;
                selectedFilters[i].filterValues = selectedFilters[i].filterValues.filter((value: IFilterVal) => {
                  return myfilterValues.some((valueFilter: IFilterVal) => {
                    return value.title !== valueFilter.title;
                  });
                });
              }
            }
          }
        }
        // remove whole object because it is last filterValue
        if (found === false) {
          selectedFilters = selectedFilters.filter((selFilter: IFilters) => {
              return selFilter.filterName !== filter.filterName;
          });
        }
        if (multi === true && myfilterValues.length > 0) {
          this.setState({
            selectedFilters: selectedFilters
          });
        } else {
          const newItems: any = this.filter(selectedFilters, this.state.valueFilter);
          this.setState({
            selectedFilters: selectedFilters,
            items: newItems,
            columns: this.columnsFilter (selectedFilters),
            filters: this.createFilters(newItems, selectedFilters)
          });
        }
      }
    }
  }
  private createFilters = (data: any, selectedFiltersInput: IFilters[]): IFilters[] => {
    this.setState({
        isLoadingData: true
    });
    const filters: IFilters[] = [];
    const filterArray: IFilterConfiguration[] = (this.props.filterConfiguration) ? this.props.filterConfiguration : [];
    let selectedFiltersInternal: IFilters[] = [];
    if (selectedFiltersInput !== undefined) {
      selectedFiltersInternal = selectedFiltersInput;
    } else {
      selectedFiltersInternal = this.state.selectedFilters;
    }
    if (selectedFiltersInternal.length > 0) {
      for (let i: number = 0; i < filterArray.length; i++) {
        // check if filter column is in columns definition for the view
        const colDef: any = this.state.columnsDefinition.filter((e) => e.RealFieldName === filterArray[i].filterName);
        if (colDef.length > 0) {
          const checkIfSelected: IFilters[] = selectedFiltersInternal.filter((selFilter: IFilters) => selFilter.filterName === filterArray[i].filterName);
          const filterValues: IFilterVal[] =
            filterArray[i].filterMode === FilterTemplateOption.FixedDateRange ?
                checkIfSelected.length > 0 ?
                    checkIfSelected[0].filterValues
                    :
                    [{title: 'null', count: 0}]
            :
            this.getFilters(data, filterArray[i], colDef[0].Type);
          if (checkIfSelected.length > 0) {
            if (filterValues.length > 0) {
              filters.push({
                filterTitle: colDef[0].DisplayName,
                filterType: colDef[0].Type,
                filterName: filterArray[i].filterName,
                filterValues: filterValues,
                filterMode: filterArray[i].filterMode,
                showExpanded: filterArray[i].showExpanded,
                showValueFilter: filterArray[i].showValueFilter,
                executeSearch: false
              });
            } else {
              filters.push(checkIfSelected[0]);
            }
          } else {
            if (filterValues.length > 0) {
              filters.push({
                filterTitle: colDef[0].DisplayName,
                filterType: colDef[0].Type,
                filterName: filterArray[i].filterName,
                filterValues: filterValues,
                filterMode: filterArray[i].filterMode,
                showExpanded: filterArray[i].showExpanded,
                showValueFilter: filterArray[i].showValueFilter,
                executeSearch: false
              });
            }
          }
        }
      }
    } else {
      if (filterArray.length > 0) {
        for (let i: number = 0; i < filterArray.length; i++) {
          // check if filter column is in columns definition for the view
          const colDef: any[] = this.state.columnsDefinition.filter((e) => e.RealFieldName === filterArray[i].filterName);
          if (colDef.length > 0) {
            const filterValues: IFilterVal[] = filterArray[i].filterMode === FilterTemplateOption.FixedDateRange ?
              [{title: 'null', count: 0}]
              :
              this.getFilters(data, filterArray[i], colDef[0].Type);
            filters.push({
                filterTitle: colDef[0].DisplayName,
                filterType: colDef[0].Type,
                filterName: filterArray[i].filterName,
                filterValues: filterValues,
                filterMode: filterArray[i].filterMode,
                showExpanded: filterArray[i].showExpanded,
                showValueFilter: filterArray[i].showValueFilter,
                executeSearch: false
            });
          }
        }
      }
    }
    this.setState({
      isLoadingData: false
    });
    return filters;
  }
  private filter = (selectedFilters: IFilters[], searchValue: string): any => {
    this.setState({
      isLoadingData: true
    });
    let newGridData: any[] = this.search(this.gridItems, searchValue);
    // we have erased all filters
    if (selectedFilters.length === 0) {
      newGridData = this.copyAndSort(newGridData, this.state.sortedColumnKey, this.state.isSortedDescending);
    } else {
      for (let i: number = 0; i < selectedFilters.length; i++) {
        if (((selectedFilters[i].filterMode !== FilterTemplateOption.FixedDateRange) &&
          (selectedFilters[i].filterMode !== FilterTemplateOption.CheckBoxMulti)) ||
          ((selectedFilters[i].filterMode === FilterTemplateOption.CheckBoxMulti ) &&
          (selectedFilters[i].executeSearch === true))) {
          let filterGridData: any[] = [];
          for (let j: number = 0; j < selectedFilters[i].filterValues.length; j++) {
            let partialGridData: any[] = [];
            if (selectedFilters[i].filterType === 'User') {
              if (selectedFilters[i].filterValues[j].title === 'N/A') {
                partialGridData = newGridData.filter((item) => {
                  return item[selectedFilters[i].filterName] === '';
                });
              } else {
                partialGridData = newGridData.filter((item) => {
                  return item[selectedFilters[i].filterName] ?
                    item[selectedFilters[i].filterName][0].title === selectedFilters[i].filterValues[j].title
                    :
                    undefined ;
                });
              }
            } else if (selectedFilters[i].filterType === 'Lookup') { //Taxonomy Lookup - be aware of connected lists (to view connected list columns change code)
              partialGridData = newGridData.filter((item) => {
                return item[selectedFilters[i].filterName] ?
                item[selectedFilters[i].filterName].Label === selectedFilters[i].filterValues[j].title
                :
                undefined ;
              });
            } else {
              partialGridData = newGridData.filter((item) => item[selectedFilters[i].filterName] === selectedFilters[i].filterValues[j].title);
            }
            filterGridData = [...filterGridData, ...partialGridData];
          }
          newGridData = filterGridData;
        } else if (selectedFilters[i].filterMode === FilterTemplateOption.FixedDateRange) {
          if (selectedFilters[i].filterValues.length > 0) {
            moment.locale(this.props.currentCultureName);
            newGridData = newGridData.filter((item) => {
              const splitDateRange: string[] = selectedFilters[i].filterValues[0].title.split('|');
              let inRangeFrom: boolean = true;
              let inRangeTo: boolean = true;
              // from
              if (splitDateRange[0] !== 'null') {
                if (item[selectedFilters[i].filterName].indexOf(':') === -1) { // date without time
                  if (moment(item[selectedFilters[i].filterName], 'L').isValid()) {
                    inRangeFrom = moment(item[selectedFilters[i].filterName], 'L').isSameOrAfter(splitDateRange[0]);
                  }
                } else {
                  if (moment(item[selectedFilters[i].filterName]).isValid()) {
                    inRangeFrom = moment(item[selectedFilters[i].filterName]).isSameOrAfter(splitDateRange[0]);
                  }
                }
              }
              // to
              if (splitDateRange[1] !== 'null') {
                if (item[selectedFilters[i].filterName].indexOf(':') === -1) { // date without time
                  if (moment(item[selectedFilters[i].filterName], 'L').isValid()) {
                    inRangeTo = moment(item[selectedFilters[i].filterName], 'L').isSameOrBefore(splitDateRange[1]);
                  }
                } else {
                  if (moment(item[selectedFilters[i].filterName]).isValid()) {
                      inRangeTo = moment(item[selectedFilters[i].filterName]).isSameOrBefore(splitDateRange[1]);
                  }
                }
              }
              if (this.props.debug) {
                console.log ('RESULT:');
              }
              if (this.props.debug) {
                console.log (inRangeTo && inRangeFrom);
              }
              return inRangeTo && inRangeFrom;
            });
          }
        }
      }
      newGridData = this.copyAndSort(newGridData, this.state.sortedColumnKey, this.state.isSortedDescending);
    }
    if (this.state.valueFilter !== searchValue) {
      this.setState ({
        filters: this.createFilters(newGridData, undefined)
      });
    }
    // update results count in command bar
    $('div[data-command-key="resultsCount"] span').text(strings.Table.ResultsCount + ' : ' + newGridData.length);
    this.setState({
      isLoadingData: false
    });
    return newGridData;
  }
  private onRemoveAllFilters = () => {
    const newItems: any = this.filter([], this.state.valueFilter);
    this.setState({
        selectedFilters: [],
        items: newItems,
        columns: this.columnsFilter ([]),
        filters: this.createFilters(newItems, [])
    });
  }
  private getSelectionDetails(): any[] {
    const selectionCount: number = this.selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return [];
      case 1:
        return this.selection.getSelection();
      default:
        return [];
    }
  }
  private onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props) {
      return undefined;
    }
    return (
      <Sticky stickyPosition={StickyPositionType.Header}>
        {defaultRender!({
          ...props
        })}
      </Sticky>
    );
  }
  private columnsFilter = (selectedFilters: IFilters[]): any => {
    const columns: IColumn[] = this.state.columns;
    columns.forEach(column => {
        const filtersColumn: IFilters[] = selectedFilters.filter ((selFilter) => selFilter.filterName === column.fieldName);
        column.isFiltered = (filtersColumn.length > 0) ? true : false;
    });
    return columns;
  }
  private onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this.getContextualMenuProps(ev, column)
      });
    }
  }
  private onColumnContextMenu = (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this.getContextualMenuProps(ev, column)
      });
    }
  }
  private getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps {
    const items: IContextualMenuItem[] = [
      {
        key: 'aToZ',
        name: 'A -> Z',
        iconProps: { iconName: 'SortUp' },
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => this.onSortColumn(column.key, false)
      },
      {
        key: 'zToA',
        name: 'Z -> A',
        iconProps: { iconName: 'SortDown' },
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => this.onSortColumn(column.key, true)
      }
    ];
    if (this.props.filterConfiguration.filter((e: IFilterConfiguration) => e.filterName === column.fieldName).length > 0 ) {
        items.push({
            key: 'filter',
            name: strings.Table.filterText,
            iconProps: { iconName: 'Filter' },
            canCheck: false,
            checked: false,
            onClick: () => this.setState({...this.state, showFiltersPanel: true, filterField: column.fieldName})
      });
    }
    return {
      items: items,
      target: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      gapSpace: 0,
      isBeakVisible: false,
      onDismiss: this.onContextualMenuDismissed
    };
  }
  private onItemContextMenu = (item: any, index: number, ev: MouseEvent): boolean => {
    const contextualMenuProps: IContextualMenuProps = {
      target: ev.target as HTMLElement,
      items: [],
      onDismiss: () => {
        this.setState({
          contextualMenuProps: undefined
        });
      }
    };
    if (index > -1) {
      this.setState({
        contextualMenuProps: contextualMenuProps
      });
    }
    return false; // to call event.preventDefault
  }
  private onSortColumn = (columnKey: string, isSortedDescending: boolean): void => {
    this.setState({
        isLoadingData: true
    });
    const sortedItems: any[] = this.copyAndSort(this.state.items, columnKey, isSortedDescending);
    this.setState({
        isLoadingData: false,
        items: sortedItems,
        columns: this.buildColumns(
            sortedItems,
            this.state.columnsDefinition,
            true,
            this.onColumnClick,
            columnKey,
            isSortedDescending,
            this.onColumnContextMenu
        ),
        isSortedDescending: isSortedDescending,
        sortedColumnKey: columnKey
    });
  }
  private onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined
    });
  }
  private onValueFilterCleared = () => {
    this.setState({
      valueFilter: '',
      items: this.filter(this.state.selectedFilters, '')
    });
  }
  private onValueFilterChanged = (newValue: string) => {
    this.setState({
      valueFilter: newValue,
      items: this.filter(this.state.selectedFilters, newValue)
    });
  }
  private onValueFilterChanged2 = (newValue: string) => {
    if (newValue === '') {
      this.setState({
        valueFilter: '',
        items: this.filter(this.state.selectedFilters, '')
      });
    }
  }
  private onValueFilterClick = (event: React.MouseEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    event.stopPropagation();
  }
  /**
   * compare function to pass to the copyAndSort function
   * @param key Column
   * @param isSortedDescending 
   */
  private compareFn = (key: string, isSortedDescending?: boolean) => {
    const debug: boolean = this.props.debug;
    return (a: any, b: any): number => {
        if (debug === true) {
            console.log (a[key]);
            console.log (b[key]);
        }
        if (a[key] === null || a[key] === undefined) return 1;
        if (b[key] === null || b[key] === undefined) return -1;

        if (typeof a[key] === 'object' && a[key] !== null && !($.isArray(a[key]))) { // taxonomy
            if (debug === true) {
              console.log ('sorting taxonomy');
            }
            if (isSortedDescending) {
                if (a[key].Label < b[key].Label) {
                    return 1;
                } else {
                    return -1;
                }
            } else {
                if (a[key].Label > b[key].Label) {
                    return 1;
                } else {
                    return -1;
                }
            }
        } else if ($.isArray(a[key])) { // people
            if (debug === true) {
              console.log ('sorting people');
            }
            if (isSortedDescending) {
                return b[key][0].title.split(' ')[0].localeCompare(a[key][0].title.split(' ')[0], 'sk');
            } else {
                return a[key][0].title.split(' ')[0].localeCompare(b[key][0].title.split(' ')[0], 'sk');
            }
        } else if (a[key].indexOf('€') !== -1) { // currency
            if (debug === true) {
              console.log ('sorting currency');
            }
            if (isSortedDescending) {
                return Number(+a[key + '.']) - Number(+b[key + '.']);
            } else {
                return Number(+b[key + '.']) - Number(+a[key + '.']);
            }
        } else if ((a[key].indexOf(':') !== -1) && (b[key].indexOf(':') !== -1) && (moment(a[key]).isValid() && moment(b[key]).isValid())) { // date with time
            if (debug === true) {
              console.log ('sorting date with time');
            }
            if (isSortedDescending) {
                return moment(b[key]).diff(moment(a[key]));
            } else {
                return moment(a[key]).diff(moment(b[key]));
            }
        } else if ((!/[a-zA-Z]/g.test(a[key])) && moment(a[key], 'L').isValid() && moment(b[key], 'L').isValid()) { // date (without time)
            if (debug === true) {
              console.log ('sorting date without time');
            }
            if (isSortedDescending) {
                return moment(b[key], 'L').diff(moment(a[key], 'L'));
            } else {
                return moment(a[key], 'L').diff(moment(b[key], 'L'));
            }
        } else {
            if (debug === true) {
              console.log ('sorting default');
            }
            if (isSortedDescending) {
                if (a[key] < b[key]) {
                    return 1;
                } else {
                    return -1;
                }
            } else {
                if (a[key] > b[key]) {
                    return 1;
                } else {
                    return -1;
                }
            }
        }
    };
  }
  private copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
      return items.slice(0).sort(this.compareFn(columnKey, isSortedDescending));
  }
  /**
   * Function to create unique values for list data and particular column. These values represents filter values.
   * Function creates IFilterVal array with unique values sorted (based on configuration) and count 
   * Warning: Lookup value can be Taxonomy or ConnectedList - for Taxonomy use .Label and for connected list investigate - i think it is [0].Value
   * @param array Whole data
   * @param field: IFilterConfiguration for which filters should be created
   * @param colDefType filed type (User, Lookup, Number, ...)
   */
  private getFilters = (array: string[], field: IFilterConfiguration, colDefType: string): IFilterVal[] => {
    const unique: number[] = [];
    const distinct: IFilterVal[] = [];
    for (let i: number = 0; i < array.length; i++ ) {
        if (colDefType === 'User') {
            if (array[i][field.filterName] === '') {
                if (!unique[array[i][field.filterName]]) {
                    distinct.push(
                        {
                            title: 'N/A',
                            count: 1
                        }
                    );
                    unique[array[i][field.filterName]] = 1;
                } else {
                    distinct[distinct.findIndex(p => p.title === 'N/A')].count += 1;
                }
            } else {
                if (!unique[array[i][field.filterName][0].title]) {
                    distinct.push(
                        {
                            title: array[i][field.filterName][0].title,
                            count: 1
                        }
                    );
                    unique[array[i][field.filterName][0].title] = 1;
                } else {
                  distinct[distinct.findIndex(p => p.title === array[i][field.filterName][0].title)].count += 1;
                }
            }
        } else if (colDefType === 'Lookup') { //Taxonomy Lookup - be aware of connected lists (to view connected list columns change code)
            if (!unique[array[i][field.filterName].Label]) {
              distinct.push(
                  {
                      title: array[i][field.filterName].Label,
                      count: 1
                  }
              );
              unique[array[i][field.filterName].Label] = 1;
            } else {
              distinct[distinct.findIndex(p => p.title === array[i][field.filterName].Label)].count += 1;
            }
        } else if (!unique[array[i][field.filterName]]) {
            distinct.push(
                {
                    title: array[i][field.filterName],
                    count: 1
                }
            );
            unique[array[i][field.filterName]] = 1;
        } else {
          distinct[distinct.findIndex(p => p.title === array[i][field.filterName])].count += 1;
        }
    }
    let result: IFilterVal[] = [];
    switch (field.filterSortType) {
        case FiltersSortOption.ByNumberOfResults:
            result = distinct.slice(0).sort((a: any, b: any) => ((field.filterSortDirection === FiltersSortDirection.Descending ? a.count < b.count : a.count > b.count) ? 1 : -1));
            break;
        default:
            result = distinct.slice(0).sort((a: any, b: any) => ((field.filterSortDirection === FiltersSortDirection.Descending ? a.title < b.title : a.title > b.title) ? 1 : -1));
            break;
    }
    return result;
  }
}
