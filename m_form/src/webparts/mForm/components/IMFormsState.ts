import { IFieldSchema } from './utilities/RenderListData';

export interface IMFormsState {
    fieldSchema: IFieldSchema[];
    originalFieldSchema: IFieldSchema[];
    contenTypeId: string;
    item: any;
    originalItem: any;
    isLoading: boolean;
    isSaving: boolean;
    editableFields: string[];
    fieldErrors: { [fieldName: string]: string };
    requiredFieldEmpty: boolean;
    spListInfoServerRelativeUrl: string;
    newAttachments: {file: any}[];
    deleteAttachments: any[];
    errors: string[];
    notifications: string[];
}