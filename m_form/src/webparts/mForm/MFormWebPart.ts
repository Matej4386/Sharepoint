import * as React from 'react';
import * as ReactDom from 'react-dom';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyPanePropertyEditor } from '@pnp/spfx-property-controls/lib/PropertyPanePropertyEditor';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
import ConfigureWebPart from './components/utilities/ConfigureWebpart/ConfigureWebPart';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  IPropertyPaneGroup,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import MForms from './components/MForms';

export interface IMFormWebPartProps {
  listId: string;
  contentTypeId: string;
  renderMethod: number;
  height: string;
  isEditableInViewForm: boolean;
  listItemId: number;
  debug: boolean;
  redirectUrl: string;
  queryStringListItemIdParam: string;
  inlineForm: boolean;
  inlineFormLabelWidth: string;
}

export default class MFormWebPart extends BaseClientSideWebPart<IMFormWebPartProps> {
  private viewModeOptions: IPropertyPaneDropdownOption[];
  public constructor() {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      this.viewModeOptions = [
        {
          key: 1,
          text: 'Display Form'
        },
        {
          key: 2,
          text: 'Edit Form'
        },
        {
          key: 3,
          text: 'New Form'
        }
      ];
    });
  }
  public render(): void {
    let element: any;
    let itemId: number;
    if (!this.properties.queryStringListItemIdParam) {
      this.properties.queryStringListItemIdParam = 'mID';
    }
    if (!this.properties.renderMethod) {
      this.properties.renderMethod = 3;
    }
    itemId = Number(this.properties.listItemId);
    if ((isNaN(itemId) || itemId === 0)) {
      // if item Id is not a number we assume it is a query string parameter
      const urlParams: URLSearchParams = new URLSearchParams(window.location.search);
      itemId = Number(urlParams.get(this.properties.queryStringListItemIdParam));
    }
    if ((isNaN(itemId) || itemId === 0)) {
      itemId = null;
    }
    if (this.properties.listId) {
      element = React.createElement(
        MForms,
        {
          debug: this.properties.debug, // debug?: boolean - console.log
          context: this.context, // current context
          listId: this.properties.listId, // list ID
          listItemId: this.properties.renderMethod !== 3 ? itemId : null , // ItemID?
          contentTypeId: this.properties.contentTypeId ? this.properties.contentTypeId : null, // content type ID if in list are more contecnt types used
          formRenderMethod: this.properties.renderMethod, // 1 - display, 2 - edit, 3 - new
          height: this.properties.height ? this.properties.height : null, // default is 650px
          isEditableInViewForm: this.properties.isEditableInViewForm,
          inlineForm: this.properties.inlineForm ? this.properties.inlineForm : false,
          inlineFormLabelWidth: this.properties.inlineFormLabelWidth ? this.properties.inlineFormLabelWidth : '10rem',
          /**
           * Function called after succesfull save (create or update)
           * @param listItem item state after succesfull save
           * @param itemId item ID (number)
           */
          onSuccesSave: (listItem, newItemId) => {
            console.log ('onSuccessSave');
            console.log (listItem);
            console.log (newItemId);
            if ((this.properties.redirectUrl) && (this.properties.renderMethod !== 1)) {
              // redirect to configured URL after successfully submitting form
              window.location.href = this.properties.redirectUrl.replace('[ID]', newItemId.toString());
            }
          },
          /**
           * Function called after error in save function (create or update)
           * @param listItem item state after error
           * @param error error object (in catch)
           */
          onErrorSave: (listItem, error) => {
            console.log ('onErrorSave');
            console.log (listItem);
            console.log (error);
          },
          onCancelForm: () => {
            console.log ('Form cancel');
          },
          onEditFieldsBeforeUpdate: (fields) => {
            console.log ('onEditFieldsBeforeUpdate');
            console.log (fields);
            return fields;
          }
        }
      );
    } else {
      // show configure web part react component
      element = React.createElement(
        ConfigureWebPart,
        {
          webPartContext: this.context,
          title: 'Form',
          description: 'Configure first',
          buttonText: 'Configure'
        }
      );
    }

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const mainGroup: IPropertyPaneGroup[] = [{
      groupName: 'Basic Settings',
        groupFields: [
          PropertyFieldListPicker('listId', {
            label: 'List',
            selectedList: this.properties.listId,
            includeHidden: false,
            orderBy: PropertyFieldListPickerOrderBy.Title,
            disabled: false,
            onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
            properties: this.properties,
            context: this.context,
            onGetErrorMessage: null,
            deferredValidationTime: 500,
            key: 'listId'
          }),
          PropertyPaneDropdown('renderMethod', {
            label: 'Form type',
            options: this.viewModeOptions,
            selectedKey: this.viewModeOptions[this.properties.renderMethod - 1].key,
            disabled: !this.properties.listId
          }),
          PropertyPaneTextField('contentTypeId', {
            label: 'Content type ID',
            disabled: !this.properties.listId
          }),
          PropertyPaneTextField('height', {
            label: 'Form height (default: 650px)',
            disabled: !this.properties.listId
          }),
          PropertyPaneCheckbox('inlineForm', {
            text: 'Inline form?'
          })
        ]
    }];
    if (this.properties.inlineForm === true) {
      mainGroup[0].groupFields.push(
        PropertyPaneTextField('inlineFormLabelWidth', {
          label: 'Inline form Label fixed width (default: 10rem)',
          disabled: !this.properties.listId
        }),
      );
    }
    if (this.properties.renderMethod === 1) { // Display Form
      mainGroup[0].groupFields.push(
        PropertyPaneTextField('listItemId', {
          label: 'List item ID (number)',
          disabled: !this.properties.listId
        }),
        PropertyPaneTextField('queryStringListItemIdParam', {
          label: 'Query string param name',
          description: 'Name of the parameter in query string [ParamName]. Example: /DispForm.aspx?[ParamName]=[ItemId]',
          disabled: !this.properties.listId
        }),
        PropertyPaneCheckbox('isEditableInViewForm', {
          text: 'Is editable in View form?',
          disabled: this.properties.renderMethod !== 1 ? true : false
        })
      );
    } else if (this.properties.renderMethod === 2) { // Edit Form
      mainGroup[0].groupFields.push(
        PropertyPaneTextField('listItemId', {
          label: 'List item ID (number)',
          disabled: !this.properties.listId
        }),
        PropertyPaneTextField('queryStringListItemIdParam', {
          label: 'Query string param name',
          description: 'Name of the parameter in query string [ParamName]. Example: /DispForm.aspx?[ParamName]=[ItemId]',
          disabled: !this.properties.listId
        }),
        PropertyPaneTextField('redirectUrl', {
          label: 'Redirect url',
          description: 'mID is required as param ID. Can contain [ID] as a placeholder to be replaced by ID of updated or created item. Example: /list/Test/DispForm.aspx?[ParamName]=[ID]',
          disabled: !this.properties.listId
        })
      );
    } else { // New Form
      mainGroup[0].groupFields.push(
        PropertyPaneTextField('redirectUrl', {
          label: 'Redirect url',
          description: '[ParamName] is required as param ID. Can contain [ID] as a placeholder to be replaced by ID of updated or created item. Example: /list/Test/DispForm.aspx?[ParamName]=[ID]',
          disabled: !this.properties.listId
        })
      );
    }
    const otherGroups: IPropertyPaneGroup[] = [
      {
        groupName: 'Debug',
        groupFields: [
          PropertyPaneCheckbox('debug', {
            text: 'Debug?'
          })
        ]
      },
      {
        groupName: 'Properties editor',
        groupFields: [
          PropertyPanePropertyEditor({
            webpart: this,
            key: 'propertyEditor'
          })
        ]
      }
    ];
    return {
      pages: [
        {
          header: {
            description: 'Configuration'
          },
          groups: [...mainGroup, ...otherGroups]
        },
        {
          groups: [{
            groupName: 'About',
            groupFields: [
              PropertyPaneWebPartInformation({
                description: `
                  <span>Author:
                    <ul style='list-style-type: none;padding: 0;'>
                      <li>Matej Jurikovič
                        <a href='https://github.com/Matej4386' style='margin-left: "1rem"'>
                          <img width='16px' src='https://github.githubassets.com/images/modules/site/sponsors/pixel-mona-heart.gif'/>
                        </a>
                      </li>
                    </ul>
                  </span>`,
                key: 'authors'
              }),
              PropertyPaneLabel('', {
                text: `Version: ${this && this.manifest.version ? this.manifest.version : ''}`
              })
            ]
          }]
        }
      ]
    };
  }
}
