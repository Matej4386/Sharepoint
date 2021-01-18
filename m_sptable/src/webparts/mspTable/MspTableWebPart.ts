import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as $ from 'jquery';
import {sp} from '@pnp/sp';
import PnPTelemetry from '@pnp/telemetry-js';
import {IFilterConfiguration} from './components/modules/datatypes/IFilterConfiguration';
import {FilterTemplateOption} from './components/modules/datatypes/FilterTemplateOption';
import {FiltersSortOption} from './components/modules/datatypes/FiltersSortOptions';
import {FiltersSortDirection} from './components/modules/datatypes/FiltersSortDirection';
import {SelectionMode} from 'office-ui-fabric-react/lib/DetailsList';
import {PropertyFieldListPicker, PropertyFieldListPickerOrderBy} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'MspTableWebPartStrings';
import MspTable from './components/MspTable';
import {IMspTableProps} from './components/IMspTableProps';

export interface IMspTableWebPartProps {
  listId: string;
  filterConfiguration: IFilterConfiguration[];
  searchConfiguration: string;
}

export default class MspTableWebPart extends BaseClientSideWebPart<IMspTableWebPartProps> {
  private _propertyPanePropertyEditor = null;
  private _propertyFieldCollectionData = null;
  private _customCollectionFieldType = null;
  
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      $('#workbenchPageContent').prop('style', 'max-width: none');
      $('.SPCanvas-canvas').prop('style', 'max-width: none');
      $('.CanvasZone').prop('style', 'max-width: none');
      sp.setup(
        {
          spfxContext: this.context,
          sp: {
            headers: {
              Accept: 'application/json; odata=nometadata'
            }
          }
        }
      );
      // Disable PnP Telemetry
      const telemetry: PnPTelemetry = PnPTelemetry.getInstance();
      if (telemetry.optOut) { telemetry.optOut(); }
    });
  }
  public render(): void {
    const element: React.ReactElement<IMspTableProps> = React.createElement(
      /**
       * Do not forget to add Locales for MspTable to config.js
       */
      MspTable,
      {
        currentCultureName: this.context.pageContext.cultureInfo.currentCultureName,
        debug: true,
        listToDisplayId: this.properties.listId,
        filterConfiguration: this.properties.filterConfiguration,
        searchConfiguration: this.properties.searchConfiguration,
        onInitSortedKey: undefined,
        selectionMode: SelectionMode.single,
        commandBar: true,
        onCommandBarChange: undefined,
        renderInfo: (error: boolean, message: string) => console.log (message)
      }
    );
    ReactDom.render(element, this.domElement);
  }
  protected async onPropertyPaneConfigurationStart() {
    await this.loadPropertyPaneResources();
  }
  protected async loadPropertyPaneResources(): Promise<void> {
    const { PropertyPanePropertyEditor } = await import(
      /* webpackChunkName: 'filter-property-pane' */
      '@pnp/spfx-property-controls/lib/PropertyPanePropertyEditor'
    );
    this._propertyPanePropertyEditor = PropertyPanePropertyEditor;
    const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
      /* webpackChunkName: 'filter-property-pane' */
      '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
    );
    this._propertyFieldCollectionData = PropertyFieldCollectionData;
    this._customCollectionFieldType = CustomCollectionFieldType;
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('listId', {
                  label: strings.PropertyPane.ListFieldLabel,
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 500,
                  key: 'listPickerFieldId'
                }),
                this._propertyFieldCollectionData('filterConfiguration', {
                  manageBtnLabel: strings.PropertyPane.EditRefinersLabel,
                  key: 'refiners',
                  enableSorting: true,
                  panelHeader: strings.PropertyPane.EditRefinersLabel,
                  panelDescription: strings.PropertyPane.RefinersFieldDescription,
                  label: strings.PropertyPane.RefinersFieldLabel,
                  value: this.properties.filterConfiguration,
                  fields: [
                    {
                      id: 'filterName',
                      title: strings.PropertyPane.Templates.FilterInternalName,
                      type: this._customCollectionFieldType.string,
                    },
                    {
                      id: 'filterMode',
                      title: strings.PropertyPane.Templates.FilterMode,
                      type: this._customCollectionFieldType.dropdown,
                      options: [
                        {
                          key: FilterTemplateOption.CheckBox,
                          text: strings.PropertyPane.Templates.RefinementItemTemplateLabel
                        },
                        {
                          key: FilterTemplateOption.CheckBoxMulti,
                          text: strings.PropertyPane.Templates.MutliValueRefinementItemTemplateLabel
                        },
                        {
                          key: FilterTemplateOption.Persona,
                          text: strings.PropertyPane.Templates.PersonaRefinementItemLabel,
                        },
                        {
                          key: FilterTemplateOption.FixedDateRange,
                          text: strings.PropertyPane.Templates.FixedDateRangeRefinementItemLabel,
                        },
                      ]
                    },
                    {
                      id: 'filterSortType',
                      title: strings.PropertyPane.Templates.RefinerSortTypeLabel,
                      type: this._customCollectionFieldType.dropdown,
                      options: [
                        {
                          key: FiltersSortOption.Default,
                          text: "--"
                        },
                        {
                          key: FiltersSortOption.ByNumberOfResults,
                          text: strings.PropertyPane.Templates.RefinerSortTypeByNumberOfResults,
                          ariaLabel: strings.PropertyPane.Templates.RefinerSortTypeByNumberOfResults
                        },
                        {
                          key: FiltersSortOption.Alphabetical,
                          text: strings.PropertyPane.Templates.RefinerSortTypeAlphabetical,
                          ariaLabel: strings.PropertyPane.Templates.RefinerSortTypeAlphabetical
                        }
                      ]
                    },
                    {
                      id: 'filterSortDirection',
                      title: strings.PropertyPane.Templates.RefinerSortTypeSortOrderLabel,
                      type: this._customCollectionFieldType.dropdown,
                      options: [
                          {
                              key: FiltersSortDirection.Ascending,
                              text: strings.PropertyPane.Templates.RefinerSortTypeSortDirectionAscending,
                              ariaLabel: strings.PropertyPane.Templates.RefinerSortTypeSortDirectionAscending
                          },
                          {
                              key: FiltersSortDirection.Descending,
                              text: strings.PropertyPane.Templates.RefinerSortTypeSortDirectionDescending,
                              ariaLabel: strings.PropertyPane.Templates.RefinerSortTypeSortDirectionDescending
                          }
                      ]
                    },
                    {
                      id: 'showExpanded',
                      title: strings.PropertyPane.ShowExpanded,
                      type: this._customCollectionFieldType.boolean
                    },
                    {
                      id: 'showValueFilter',
                      title: strings.PropertyPane.showValueFilter,
                      type: this._customCollectionFieldType.boolean
                    }
                  ]
                }),
                PropertyPaneTextField('searchConfiguration', {
                  label: strings.PropertyPane.Search
                })
              ]
            },
            {
              groupName: strings.PropertyPane.debugTitle,
              groupFields: [
                PropertyPaneCheckbox('debug', {
                  text: strings.PropertyPane.debugLabel
                })
              ]
            },
            {
              groupName: strings.PropertyPane.PropertyEdit,
              groupFields: [
                this._propertyPanePropertyEditor({
                  webpart: this,
                  key: 'propertyEditor'
                })    
              ]
            }
          ]
        }
      ]
    };
  }
}
