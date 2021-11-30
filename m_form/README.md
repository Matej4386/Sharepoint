# M-FORM

SPFx Webpart component to render SharePoint list forms (Display, Edit, New) for SharePoint 2019 on-premise and SharePoint online. 

#### SPFx Framework version:
##### ~1.4.0

#### Supports: 
- inline editting in Display form, 
- 2 rendering types: normal and inline (label and input are inline), 
- configurable via webpart properties with hooks (valueUpdated, beforeSave, ...), 
- it uses AddValidateUpdateItemUsingPath and ValidateUpdateListItem calls - columns conditions are supported (for example number columns must be between 10 and 100)
- Append only fields

#### Supported fields:
- Text
- Note (Simple, RichText, AppendOnly - with previous text render)
- User and UserMulti
- Boolean
- DateTime (Date + Date and Time)
- Choice and MultiChoice
- Number
- Currency
- Attachments (Add, Delete)
- TaxonomyFieldType and TaxonomyFieldTypeMulti
- Lookup and LookupMulti
- URL (picture, link)

#### Screenshots:

![newForm](https://user-images.githubusercontent.com/58258968/144003869-5bc114a0-0631-4cf0-9067-aa3c2768e697.jpg)
![displayForm](https://user-images.githubusercontent.com/58258968/144003896-6a97f6a9-eb82-45b2-811b-086fed9960df.jpg)
![inlineForm](https://user-images.githubusercontent.com/58258968/144003914-9f307d2e-ec87-4f43-8004-bf873297344f.jpg)

## Properties
```typescript
    /**
     * debug on - console.log
     */
    debug?: boolean;
    /**
    * Current context
    */
    context: WebPartContext;
    /**
    * List id
    */
    listId: string;
    /**
    * ID of the list item to display on the form
    */
    listItemId?: number;
    /**
     * Content type id of the item
     */
    contentTypeId?: string;
    /**
     * Initial render method
     */
    formRenderMethod: number; // 1 - display, 2 - edit, 3 - new
    /**
     * Is editable in view form?
     */
    isEditableInViewForm: boolean;
    /**
     * Heigth for form - default is 650px
     */
    height?: string; // 'example 600px'
    /**
     * If the form will be inline Label + Input
     */
    inlineForm: boolean;
    /**
     * Width of label in case of Inline form to render same width for all labels
     */
    inlineFormLabelWidth: string;
    /**
     * Function called after succesfull save
     */
    onSuccesSave?: (listItem: any, listItemId: number) => void;
    /**
     * Function called after error in save
     */
    onErrorSave?: (listItem: any, error: any) => void;
    /**
     * Function to change loaded item (for example - custom default value)
     */
    onLoadChangeItem?: (item: any) => any;
    /**
     * Function to change loaded item schema (for example - custom render order)
     */
    onLoadChangeFieldSchema?: (fieldSchema: any) => any;
    /**
     * Function on cancel form
     */
    onCancelForm?: () => void;
    /**
     * Function to custom check form fields
     * ! If function is set it must return defined object !
     * FieldErrors should be update like this - fieldErrors can have already Required errors:
     * fieldErrors = {
         ...fieldErrors,
         [FieldInternalName]: 'Error description'
       };
     */
    onCustomCheckBeforeUpdate?: (stateData: any, fieldsSchema: any, fieldErrors: {[fieldName: string]: string; }) => {
        isError: boolean, // If update should continue
        errorText: string; // error text to render in Error message
        fieldErrors: {[fieldName: string]: string; } // error text under the field
    };
    /**
    * Function to change fields before UPDATE
    */
    onEditFieldsBeforeUpdate?: (fields: {
        FieldName: string;
        FieldValue: any;
        HasException: boolean;
        ErrorMessage: string;
    }[]) => {
        FieldName: string;
        FieldValue: any;
        HasException: boolean;
        ErrorMessage: string;
    }[];
    /**
     * Function to change field schema or changed value after OnValueChange item was called
     * It will update state based on return value: IMFormState
     * It must return state value!
     */
     onValueChangedHook?: (
         newValue: any,
         fieldInternalName: string,
         fieldErrors: { [fieldName: string]: string },
         currentFieldSchema: IFieldSchema[],
         originalFieldSchema: IFieldSchema[]
    ) => IMFormsState
```

#### Usage with Document Sets (files):
In src\webparts\mForm\components\utilities\FormHelpers.ts createItem function delete comment UnderlyingObjectType: 1
```typescript
public createItem (webUrl: string, listUrl: string, formValues: any): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            const httpClientOptions: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-type': 'application/json;odata=verbose',
                    'X-SP-REQUESTRESOURCES': 'listUrl=' + encodeURIComponent(listUrl),
                    'odata-version': ''
                },
                body: JSON.stringify({
                    listItemCreateInfo: {
                        __metadata: { type: 'SP.ListItemCreationInformationUsingPath' },
                        FolderPath: {
                            __metadata: { type: 'SP.ResourcePath' },
                            DecodedUrl: listUrl
                        }
                        // UnderlyingObjectType: 1 // only for files or Document sets
                    },
                    formValues,
                    bNewDocumentUpdate: false,
                    checkInComment: null
                })
            };
            const endpoint: string = `${webUrl}/_api/web/GetList(@listUrl)/AddValidateUpdateItemUsingPath`
                + `?@listUrl=${encodeURIComponent('\'' + listUrl + '\'')}`;
            this.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, httpClientOptions)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        reject(this.getErrorMessage(webUrl, response));
                    }
                })
                .then((respData) => {
                    resolve(respData.d.AddValidateUpdateItemUsingPath.results);
                })
                .catch((error) => {
                    reject(this.getErrorMessage(webUrl, error));
                });
        });
    }
```

## Contributing
Please open an issue first to discuss what you would like to change. Suggestions and changes are welcome.

## License
[MIT](https://github.com/Matej4386/Sharepoint/blob/master/LICENSE)
