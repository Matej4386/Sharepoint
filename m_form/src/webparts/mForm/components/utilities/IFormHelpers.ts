import { IFieldSchema } from './RenderListData';

export interface IFormHelpers {
    GetFormValues: (fieldsSchema: IFieldSchema[], data: any, originalData: any) => Array<{ FieldName: string, FieldValue: any, HasException: boolean, ErrorMessage: string }>;
    getDataForForm: (webUrl: string, listUrl: string, itemId: number) => Promise<any>;
    createItem: (webUrl: string, listUrl: string, formValues: any) => Promise<any>;
    updateItem: (webUrl: string, listUrl: string, formValues: any, id: string) => Promise<any>;
}
