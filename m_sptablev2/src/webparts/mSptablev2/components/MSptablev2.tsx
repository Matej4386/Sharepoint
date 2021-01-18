import * as React from 'react';
import styles from './MSptablev2.module.scss';
import * as strings from 'MspTableStrings';
import { IMSptablev2Props } from './IMSptablev2Props';
import { IMSptablev2State } from './IMSptablev2State';

import {IFilterConfiguration} from './modules/datatypes/IFilterConfiguration';
import FilterPanel from './modules/filterPanel/FilterPanel';

import * as $ from 'jquery';
import {sp} from '@pnp/sp';
import '@pnp/polyfill-ie11';
/**
 * A utility for converting HTML strings into React components.
 * Avoids the use of dangerouslySetInnerHTML and converts standard HTML elements, attributes and inline styles into their React equivalents.
 */
import ReactHtmlParser from 'react-html-parser';

import {CommandBar} from 'office-ui-fabric-react/lib/CommandBar';
import {Sticky, StickyPositionType} from 'office-ui-fabric-react/lib/Sticky';
import {Link} from 'office-ui-fabric-react/lib/Link';
import {ScrollablePane} from 'office-ui-fabric-react/lib/ScrollablePane';
import {IRenderFunction} from 'office-ui-fabric-react/lib/Utilities';
import {Image} from 'office-ui-fabric-react/lib/Image';
import {Persona, PersonaSize} from 'office-ui-fabric-react/lib/Persona';
import {IContextualMenuProps, IContextualMenuItem, DirectionalHint, ContextualMenu} from 'office-ui-fabric-react/lib/ContextualMenu';
import {CheckboxVisibility, IDetailsHeaderProps, ColumnActionsMode, IDetailsRowProps, ConstrainMode, DetailsListLayoutMode, IColumn, Selection, buildColumns} from 'office-ui-fabric-react/lib/DetailsList';
import {ShimmeredDetailsList} from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { IFilter } from './modules/datatypes/IFilter';

const itemsPreloaded: number = 100;

export default class MSptablev2 extends React.Component<IMSptablev2Props, IMSptablev2State> {
  private gridItems: any = undefined;
  private selection: Selection;
  private totalCount: number = undefined;
  private rowLimit: number = undefined;
  private hitEndLoadMore: boolean = false;
  private paginationNextHref: string[] = undefined;
  private selectedFilters: IFilter[] = undefined;
  // List relative URL
  private listUrl: string = undefined;
  constructor(props: IMSptablev2Props) {
    super(props);

    this.selection = new Selection({
        onSelectionChanged: this.onItemsSelectionChanged
    });
    this.selection.setItems([], false);
    this.state = {
        items: [],
        currentViewId: undefined,
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
        contextualMenuProps: undefined,
        sortedColumnKey: undefined,
        isSortedDescending: false,
        isLoadingData: true,
        selectionDetails: [],
        commandBarItems: [],
        farItems: [],
        nrPreloadingIterations: 0,
        showFiltersPanel: false,
        filterField: 'all'
    };
  }
  public componentDidUpdate (prevProps: IMSptablev2Props, prevState: IMSptablev2State): void {
    if (this.props.debug === true) {
      console.log ('----------- MSptablev2------------');
      console.log ('Action: ComponentDidUpdate');
      console.log ('Props:');
      console.log (this.props);
      console.log ('State:');
      console.log (this.state);
    }
    // view change
    if (prevState.currentViewId !== undefined) {
      if (prevState.currentViewId !== this.state.currentViewId) {
        if (this.props.debug === true) {
          console.log ('----------- MSptablev2------------');
          console.log ('Action: ComponentDidUpdate - currentView CHANGE (mounting data)');
        }
        this.mountData();
      }
    }
    // we should load more items
    if ((this.state.nrPreloadingIterations > 0) && (this.paginationNextHref !== undefined)) {
      if (this.props.debug === true) {
        console.log ('----------- MSptablev2------------');
        console.log ('Action: ComponentDidUpdate - Loading more data');
      }
      this.loadMoreItems();
    } else {
      if (this.state.isLoadingData === false) {
        // we loaded or items to be preloaded or hit the end
        this.hitEndLoadMore = false;
      }
    }
  }
  public componentWillMount (): void {
    this.mountData();
  }
  public render(): React.ReactElement<IMSptablev2Props> {
    let renderGrid: JSX.Element;
    let renderCommandBar: JSX.Element;
    let renderFilterpanel: JSX.Element = null;

    renderCommandBar =
      <div className={styles.CommandBarcontainerMaxWidth}>
          <CommandBar
            isSearchBoxVisible={false}
            items={this.state.commandBarItems}
            farItems={this.state.farItems}
          />
        </div>;
        // https://github.com/microsoft/fluentui/pull/9930/files !!!!!!!!!!!!! readme for info
        /**
         * Footer - do not use - more problems than solutions. For OUFR 5 it is nightmare - scrolling problems, footer hides horizontal scrollbar
         */
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
          onRenderCustomPlaceholder={this.onRenderCustomPlaceholder}
          selectionZoneProps={{
              selection: this.selection,
              disableAutoSelectOnInputElements: true,
              selectionMode: this.props.selectionMode
          }}
        />
      </div>;
    if (this.props.filterConfiguration) {
      if ((this.props.filterConfiguration.length > 0) && (this.listUrl !== undefined) && (this.state.columnsDefinition.length > 0)) {
        renderFilterpanel =
        <FilterPanel
          showPanel={this.state.showFiltersPanel}
          webAbsoluteUrl={this.props.webAbsoluteUrl}
          currentViewId={this.state.currentViewId}
          listUrl={this.listUrl}
          spHttpClient={this.props.spHttpClient}
          onUpdateShow={(show) => this.setState({showFiltersPanel: show})}
          debug={this.props.debug}
          renderInfo={this.props.renderInfo}
          filtersConfiguration={this.props.filterConfiguration}
          onFilter={this.onFilter}
          filterField={this.state.filterField}
          columnsDefinition={this.state.columnsDefinition}
        />;
      }
    }
    return(
      <div>
        {renderCommandBar}
        {renderFilterpanel}
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
  private onRenderCustomPlaceholder = (
    rowProps: IDetailsRowProps,
    index: number,
    defaultRender: (props: IDetailsRowProps) => React.ReactNode
  ): React.ReactNode => {
    // only if we stop preloading items
    if (this.hitEndLoadMore === false) {
      // to not hit double loading
      this.hitEndLoadMore = true;
      if (this.state.nrPreloadingIterations <= 0) {
        this.setState({
          // if this.rowLimit > itemsPreloaded - nrOfIterations should be at least 1
          nrPreloadingIterations: Math.floor(itemsPreloaded / this.rowLimit) === 0 ? 1 : Math.floor(itemsPreloaded / this.rowLimit)
        });
      }
    }
    return defaultRender(rowProps);
  }
  private mountData = async (): Promise<any> => {
    try {
      if (this.props.listToDisplayId !== undefined) {
        let rawData: any = undefined;
        if (this.state.currentViewId !== undefined) {
          rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream({
            RenderOptions: 1687
          },
          {
            View: this.state.currentViewId
          });
        } else {
          rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream({
            RenderOptions: 1687
          });
        }
        if (this.props.debug === true) {
          console.log ('----------- MSptablev2------------');
          console.log ('Action: MountData - data loaded');
          console.log ('RawData:');
          console.log (rawData);
        }
        if ((rawData) && (rawData.ListSchema)) {
          this.totalCount = Number(rawData.ListSchema.ItemCount);
          this.rowLimit = Number(rawData.ListData.RowLimit);
          this.listUrl = rawData.listUrlDir;
          // check if we need pagination - if there is rawData.ListData.NextHref: string we need pagination
          if (Object.prototype.hasOwnProperty.call(rawData.ListData, 'NextHref') === true) {
            // NextHref: "?Paged=TRUE&p_ID=30&PageFirstRow=31&View=32be729a-8bdd-4791-b3a2-86904c734294"
            // or enything else like this "?Paged=TRUE&p_number=654%2e000000000000&p_ID=448&PageFirstRow=51&View=32be729a-8bdd-4791-b3a2-86904c734294"
            const splitByEnd: string[] = rawData.ListData.NextHref.split('&');
            splitByEnd[0] = splitByEnd[0].substring(1); // get rid of ?
            this.paginationNextHref = splitByEnd;
          } else {
            this.paginationNextHref = undefined;
          }

          const farCommandBarItems: IContextualMenuItem[] = [];
          if (this.props.filterConfiguration) {
            if (this.props.filterConfiguration.length > 0) {
              farCommandBarItems.push({
                key: 'filters',
                name: strings.Table.filterText,
                iconProps: { iconName: 'Filter' },
                onClick: () => this.setState({...this.state, showFiltersPanel: true, filterField: 'all'})
              });
            }
          }

          // views
          const viewData: any = await sp.web.lists.getById(this.props.listToDisplayId).views.filter('Hidden ne true').get();
          if (this.props.debug === true) { console.log (viewData); }
          // if we have more than default view for the list render option to change the view
          if (viewData.length > 1) {
            const submenu: IContextualMenuItem[] = [];
            for (let i: number = 0; i < viewData.length; i++) {
              submenu.push({
                key: i.toString(),
                name: viewData[i].Title,
                canCheck: true,
                checked: rawData.ViewMetadata.Id === viewData[i].Id,
                onClick: () => {
                  this.setState({
                    currentViewId: viewData[i].Id
                  });
                }
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
          // everything set up - update state
          this.gridItems = (rawData) ?
              rawData.ListData.Row
              :
              [];
          this.setState({
            ...this.state,
            isLoadingData: false,
            // nr of iterations to load more items (nr. itemsPreloaded - const)
            nrPreloadingIterations: this.paginationNextHref === undefined ? 0 : Math.floor(itemsPreloaded / Number(this.rowLimit)),
            currentViewId: this.state.currentViewId === undefined ? rawData.ViewMetadata.Id : this.state.currentViewId,
            items: (rawData) ?
              this.paginationNextHref ?
                [...rawData.ListData.Row, ...Array(1)]
                :
                rawData.ListData.Row
              :
              [],
            columnsDefinition: (rawData) ?
              rawData.ListSchema.Field
              :
              [],
            columns: (rawData) ?
              this.buildColumns(rawData.ListData.Row, rawData.ListSchema.Field, true, this.onColumnClick, this.state.sortedColumnKey, this.state.isSortedDescending)
              :
              [],
            farItems: farCommandBarItems,
            filterField: 'reload'
          });
        }
      }
    } catch (error) {
      const errorText: string = `MspTable (mountData) -> ${strings.Errors.ErrorLoadingData}: ${error}`;
      this.setState({ ...this.state, items: {}, isLoadingData: false });
      this.props.renderInfo (true, errorText);
    }
  }
  private loadFilteredData = async (typeMap: Map<string, string>): Promise<any> => {
    try {
      if (this.props.listToDisplayId !== undefined) {
        let rawData: any = undefined;
        // to disable on shimmer load more data
        this.hitEndLoadMore = true;
        this.setState({
          ...this.state,
          isLoadingData: true,
          nrPreloadingIterations: 0,
          items: []
        });
        if (this.state.currentViewId !== undefined) {
          rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream({
            RenderOptions: 1687
          },
          {
            View: this.state.currentViewId
          },
          typeMap);
        } else {
          rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream({
            RenderOptions: 1687
          },
          {},
          typeMap);
        }
        if ((rawData) && (rawData.ListSchema)) {
          // check if we need pagination - if there is rawData.ListData.NextHref: string we need pagination
          if (Object.prototype.hasOwnProperty.call(rawData.ListData, 'NextHref') === true) {
            // NextHref: "?Paged=TRUE&p_ID=30&PageFirstRow=31&View=32be729a-8bdd-4791-b3a2-86904c734294"
            // or enything else like this "?Paged=TRUE&p_number=654%2e000000000000&p_ID=448&PageFirstRow=51&View=32be729a-8bdd-4791-b3a2-86904c734294"
            const splitByEnd: string[] = rawData.ListData.NextHref.split('&');
            splitByEnd[0] = splitByEnd[0].substring(1); // get rid of ?
            this.paginationNextHref = splitByEnd;
          } else {
            this.paginationNextHref = undefined;
          }

          // everything set up - update state
          this.gridItems = (rawData) ?
              rawData.ListData.Row
              :
              [];
          this.setState({
            ...this.state,
            isLoadingData: false,
            // nr of iterations to load more items (nr. itemsPreloaded - const)
            nrPreloadingIterations: this.paginationNextHref === undefined ? 0 : Math.floor(itemsPreloaded / Number(this.rowLimit)),
            currentViewId: this.state.currentViewId === undefined ? rawData.ViewMetadata.Id : this.state.currentViewId,
            columns: this.columnsFilter(), // to render filter icon
            items: (rawData) ?
              this.paginationNextHref ?
                [...rawData.ListData.Row, ...Array(1)]
                :
                rawData.ListData.Row
              :
              []
          });
        }
      }
    } catch (error) {
      const errorText: string = `MspTable (loadFilteredData) -> ${strings.Errors.ErrorLoadingData}: ${error}`;
      this.setState({ ...this.state, items: {}, isLoadingData: false });
      this.props.renderInfo (true, errorText);
    }
  }
  private loadMoreItems = async (): Promise<any> => {
    try {
      let rawData: any = undefined;
      const typeMap: Map<string, string> = new Map<string, string>();
      for (let i: number = 0; i < this.paginationNextHref.length; i++) {
        const split: string[] = this.paginationNextHref[i].split('=');
        typeMap.set(split[0], split[1]);
      }
      rawData = await sp.web.lists.getById(this.props.listToDisplayId).renderListDataAsStream({
        RenderOptions: 2
        },
        {
          View: this.state.currentViewId
        },
        typeMap
      );
      // check if we need pagination - if there is rawData.ListData.NextHref: string we need pagination
      if (Object.prototype.hasOwnProperty.call(rawData, 'NextHref') === true) {
        // NextHref: "?Paged=TRUE&p_ID=30&PageFirstRow=31&View=32be729a-8bdd-4791-b3a2-86904c734294"
        // or enything else like this "?Paged=TRUE&p_number=654%2e000000000000&p_ID=448&PageFirstRow=51&View=32be729a-8bdd-4791-b3a2-86904c734294"
        const splitByEnd: string[] = rawData.NextHref.split('&');
        splitByEnd[0] = splitByEnd[0].substring(1); // get rid of ?
        this.paginationNextHref = splitByEnd;
      } else {
        this.paginationNextHref = undefined; // end of list
      }
      const oldGridItems: any = [...this.gridItems];
      this.gridItems = (rawData) ?
          [...this.gridItems, ...rawData.Row]
          :
          [...this.gridItems];
      this.setState({
        ...this.state,
        nrPreloadingIterations: this.state.nrPreloadingIterations - 1,
        items: (rawData) ?
          this.paginationNextHref ?
            [...oldGridItems, ...rawData.Row, ...Array(1)]
            :
            [...oldGridItems, ...rawData.Row]
          :
          [...this.state.items]
      });
    } catch (error) {
      const errorText: string = `MspTable (loadMoreItems) -> ${strings.Errors.ErrorLoadingData}: ${error}`;
      this.props.renderInfo (true, errorText);
    }
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
    const columns: IColumn[] = buildColumns(
        items,
        canResizeColumns,
        onColumnClick,
        sortedColumnKey,
        isSortedDescending
    );
    const gridColumns: IColumn[] = [];
    columns.forEach(column => {
      const columnDef: any = columnsDefinition.filter((columnDefTemp) => column.fieldName === columnDefTemp.RealFieldName)[0];
      if (columnDef !== undefined) {
        column.maxWidth = 200;
        column.name = columnDef.DisplayName;
        // base rendering override for different columns types
        switch (columnDef.Type) {
          case 'Note':
            column.isMultiline = true;
            column.onRender = (item: any) => (
                ReactHtmlParser(item[column.fieldName])
            );
            break;
          case 'Lookup':
            switch (columnDef.FieldType) {
              // taxonomy filed type
              case 'TaxonomyFieldType':
                column.onRender = (item: any) => (
                  item[column.fieldName].Label
                );
                break;
              default:
                // standard lookup for list
                // suppports only single value
                if (Object.prototype.hasOwnProperty.call(columnDef, 'RenderAsText') === true) {
                  // use for additional lookup values renders as Text
                  if (columnDef.RenderAsText === 'TRUE') {
                    column.onRender = (item: any) => (
                      item[column.fieldName]
                    );
                  }
                } else {
                  column.onRender = (item: any) => (
                    item[column.fieldName][0].lookupValue
                  );
                }
                break;
            }
            break;
          case 'User':
            column.maxWidth = 180;
            column.onRender = (item: any) => (
                <Persona
                    text={item[column.fieldName] ? item[column.fieldName][0].title : 'N/A'}
                    size={ PersonaSize.size10 }
                />
            );
            break;
          case 'URL':
            switch (columnDef.Format) {
              case 'Hyperlink':
                column.onRender = (item: any) => (
                  <Link
                    href={item[column.fieldName] ? item[column.fieldName] : ''}
                  >
                    {item[column.fieldName + '.desc'] ? item[column.fieldName + '.desc'] : ''}
                  </Link>
                );
                break;
              case 'Image':
                column.onRender = (item: any) => (
                  <Image
                    src= {item[column.fieldName] ? item[column.fieldName] : ''}
                    alt={item[column.fieldName + '.desc'] ? item[column.fieldName + '.desc'] : ''}
                  />
                );
                break;
            }
            break;
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
  private onItemInvoked = (item: any, index: number): void => {
    /**
     * Place the code for row double click here or pass function to props
     */
    console.log ('Item invoked');
    console.log (item);
    console.log (index);
  }
  private onColumnContextMenu = (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this.getContextualMenuProps(ev, column)
      });
    }
  }
  private onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
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
    if (this.props.filterConfiguration) {
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
  private onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined
    });
  }
  private columnsFilter = (): any => {
    const columns: IColumn[] = this.state.columns;
    columns.forEach(column => {
        const filtersColumn: IFilter[] = this.selectedFilters.filter ((selFilter) => selFilter.fieldName === column.fieldName);
        const isExecuted = filtersColumn[0] ? filtersColumn[0].options.filter((options) => options.executed === true) : [];
        column.isFiltered = (isExecuted.length > 0) ? true : false;
    });
    return columns;
  }
  private onSortColumn = (columnKey: string, isSortedDescending: boolean): void => {
    /*
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
    });*/
  }
  private onItemsSelectionChanged = (): void => {
    this.setState({ selectionDetails: this.getSelectionDetails() });
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
  private handleScroll = (event: any): void => {
    const element: Element = document.querySelector("[class*='stickyAbove-']");
    if (element) {
      element.scrollLeft = event.target.scrollLeft;
    }
  }
  private onFilter = (filters: IFilter[]): void => {
    const filterTypeMap: Map<string, string> = new Map<string, string>();
    let numberOfFilters: number = 0;
    const filterField: string = 'FilterField';
    const filterValue: string = 'FilterValue';
    this.selectedFilters = [...filters];

    for (let i: number = 0; i < filters.length; i++) {
      const activeOptions = filters[i].options.filter(option => option.selected === true);
      if (activeOptions.length > 0) {
        numberOfFilters++;
        const ff: string = `${numberOfFilters === 1 ? '&' : ''}${filterField}${activeOptions.length > 1 ? 's' : ''}${numberOfFilters.toString()}`;
        const fv: string = `${filterValue}${activeOptions.length > 1 ? 's' : ''}${numberOfFilters.toString()}`;

        filterTypeMap.set(ff, filters[i].fieldName);
        if (activeOptions.length > 1) {
          let value: string = '';
          for (let j: number = 0; j < activeOptions.length; j++) {
            if (j === 0) {
              value = activeOptions[j].value;
            } else {
              value += ';#' + activeOptions[j].value;
            }
            filterTypeMap.set(fv, encodeURIComponent(value));
          }
        } else {
          filterTypeMap.set(fv, encodeURIComponent(activeOptions[0].value));
        }
      }
    }

    this.loadFilteredData(filterTypeMap);
  }

}
