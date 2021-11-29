import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IMFormsState } from './IMFormsState';
import { IFieldSchema } from './utilities/RenderListData';

export interface IMFormsProps {
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
     * If the form will be inline Label + input
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
     * Function on cacnel form
     */
    onCancelForm?: () => void;
    /**
     * Function to custom check form fields
     * !Is function is set it must return defined object!
     * FieldErrors should be update like this - fieldErrors can have alredy Required errors:
     * fieldErrors = {
                        ...fieldErrors,
                        [FieldInternalName]: 'Erro description
                    };
     */
    onCustomCheckBeforeUpdate?: (stateData: any, fieldsSchema: any, fieldErrors: {[fieldName: string]: string; }) => {
        isError: boolean, // If update should continue
        errorText: string; // error text to render in Info
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
     * It must return value
     */
     onValueChangedHook?: (
         newValue: any,
         fieldInternalName: string,
         fieldErrors: { [fieldName: string]: string },
         currentFieldSchema: IFieldSchema[],
         originalFieldSchema: IFieldSchema[]
    ) => IMFormsState
}