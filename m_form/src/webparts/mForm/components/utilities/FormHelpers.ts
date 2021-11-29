import { IFieldSchema } from './RenderListData';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IFormHelpers } from './IFormHelpers';
import { Text } from '@microsoft/sp-core-library';
import * as strings from 'MFormWebPartStrings';
import * as moment from 'moment';
import { Locales } from './Locales';

export class FormHelpers implements IFormHelpers {
    private spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient) {
        this.spHttpClient = spHttpClient;
    }

    public GetFormValues(fieldsSchema: IFieldSchema[], data: any, originalData: any)
    : Array<{ FieldName: string, FieldValue: any, HasException: boolean, ErrorMessage: string }> {
        const fields: IFieldSchema[] = fieldsSchema.filter((field) => (
                (!field.ReadOnlyField)
                && (field.InternalName in data)
                && (data[field.InternalName] !== null)
                && (data[field.InternalName] !== originalData[field.InternalName] )
            ));
        const fieldsWithVal: {
            ErrorMessage: any;
            FieldName: string;
            FieldValue: any;
            HasException: boolean;
        }[] = fields.map((field) => {
            const fieldTitle: string = field.InternalName;
            let fieldValue: any = data[field.InternalName];
            if (field.Type === 'User') {
                fieldValue = data[field.InternalName];
                const personObj: any[] = [];
                for (let i: number = 0; i < fieldValue.length; i++) {
                    fieldValue[i].Key = fieldValue[i].loginName;
                    personObj.push(fieldValue[i]);
                }
                fieldValue = JSON.stringify(personObj);
            }
            if (field.Type === 'DateTime') {
                if (fieldValue !== '') {
                    const locale: any = Locales[field.LocaleId];
                    moment.locale(locale);
                    fieldValue = moment(fieldValue).format('L HH:mm');
                    // fieldValue = moment(fieldValue).toDate().toLocaleDateString(locale).replace(/[^ -~]/g,'');
                    // fieldValue = moment(fieldValue).format('DD.MM.YYYY HH:MM');
                } else {
                    fieldValue = null;
                }
            }
            if ((field.FieldType === 'TaxonomyFieldType') || (field.FieldType === 'TaxonomyFieldTypeMulti')) {
                fieldValue = data[field.InternalName];
                if (fieldValue !== undefined) {
                    if (Array.isArray(fieldValue)) {
                        let newValue: string = '';
                        for (let i: number = 0; i < fieldValue.length; i++) {
                            newValue += fieldValue[i].name + '|' + fieldValue[i].key + ';';
                        }
                        fieldValue = newValue;
                    }
                } else {
                    fieldValue = '';
                }
            }
            if (field.FieldType === 'URL') {
                const Value: {
                    URL: string,
                    Desc: string
                } = data[field.InternalName];
                fieldValue = Value.URL + ', ' + Value.Desc;
            }
            if (field.FieldType === 'MultiChoice') {
                fieldValue = data[field.InternalName];
                fieldValue = fieldValue.join(';#');
            }
            if (field.FieldType === 'LookupMulti') {
                const values: any[] = [];
                const splitArray: any[] = data[field.InternalName].split(';#_');
                let val: any;
                for (let i: number = 0; i < splitArray.length; i++) {
                    const split: any[] = splitArray[i].split(';#');
                    if (split.length === 2) {
                        val = {
                            key: Number(split[0]),
                            text: split[1]
                        };
                        values.push(val);
                    }
                }
                let fields1: string = '';
                for (let i: number = 0; i < values.length; i++) {
                    fields1 += ';#' + values[i].key + ';#_';
                }
                fields1 += ';#';
                fieldValue = fields1;
            }
            return {
                ErrorMessage: null,
                FieldName: fieldTitle,
                FieldValue: fieldValue,
                HasException: false
            };
        });
        return fieldsWithVal;
    }

    /**
     * Retrieves the data for a specified SharePoint list form.
     *
     * @param webUrl The absolute Url to the SharePoint site.
     * @param listUrl The server-relative Url to the SharePoint list.
     * @param itemId The ID of the list item to be updated.
     * @param formType The type of form (Display, New, Edit)
     * @returns Promise representing an object containing all the field values for the list item.
     */
    public getDataForForm(webUrl: string, listUrl: string, itemId: number): Promise<any> {
        if (!listUrl || (!itemId) || (itemId === 0)) {
            return Promise.resolve({}); // no data, so returns empty
        }
        return new Promise<any>((resolve, reject) => {
            const httpClientOptions: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-type': 'application/json;odata=verbose',
                    'X-SP-REQUESTRESOURCES': 'listUrl=' + encodeURIComponent(listUrl),
                    'odata-version': ''
                }
            };
            const endpoint: string = `${webUrl}/_api/web/GetList(@listUrl)/RenderExtendedListFormData`
                + `(itemId=${itemId},formId='editform',mode='2',options=47,cutoffVersion=1)`
                + `?@listUrl=${encodeURIComponent('\'' + listUrl + '\'')}`;
            this.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, httpClientOptions)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        reject(this.getErrorMessage(webUrl, response));
                    }
                })
                .then((data) => {
                    const extendedData: any = JSON.parse(data.d.RenderExtendedListFormData);
                    const newdata: any = extendedData.Data.Row[0];
                    newdata.Attachments = extendedData.ListData.Attachments;
                    // resolve(extendedData.Data.Row[0]);
                    resolve (newdata);
                })
                .catch((error) => {
                    reject(this.getErrorMessage(webUrl, error));
                });
        });
    }
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
    public updateItem (webUrl: string, listUrl: string, formValues: any, id: string): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            const httpClientOptions: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-type': 'application/json;odata=verbose',
                    'X-SP-REQUESTRESOURCES': 'listUrl=' + encodeURIComponent(listUrl),
                    'odata-version': ''
                },
                body: JSON.stringify({
                    formValues,
                    bNewDocumentUpdate: false,
                    checkInComment: null
                })
            };
            const endpoint: string = `${webUrl}/_api/web/GetList(@listUrl)/items(@id)/ValidateUpdateListItem()`
                + `?@listUrl=${encodeURIComponent('\'' + listUrl + '\'')}`
                + `&@id=${encodeURIComponent('\'' + id + '\'')}`;
            this.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, httpClientOptions)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        reject(this.getErrorMessage(webUrl, response));
                    }
                })
                .then((respData) => {
                    resolve(respData.d.ValidateUpdateListItem.results);
                })
                .catch((error) => {
                    reject(this.getErrorMessage(webUrl, error));
                });
        });
    }
    /**
     * Returns an error message based on the specified error object
     * @param error : An error string/object
     */
    private getErrorMessage(webUrl: string, error: any): string {
        console.log (error);
        let errorMessage: string = error.statusText ?
            error.statusText
            :
            error.statusMessage ?
                error.statusMessage
                :
                error;
        const webServerRelativeUrl: string = webUrl

        if (error.status === 403) {
            errorMessage = Text.format(strings.Errors.ErrorWebAccessDenied, webServerRelativeUrl);
        } else if (error.status === 404) {
            errorMessage = Text.format(strings.Errors.ErrorWebNotFound, webServerRelativeUrl);
        }
        return errorMessage;
    }
}