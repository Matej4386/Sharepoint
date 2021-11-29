import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IFieldSchema } from '../utilities/RenderListData';

export interface IFieldsProps {
    debug: boolean;
    context: WebPartContext;
    field: IFieldSchema;
    value: any;
    originalValue: any;
    errorMessage: string;
    isSaving: boolean;
    isEditableInViewForm: boolean;
    formRenderMethod: number; // 1 - display, 2 - edit, 3 - new
    baseRenderMethod: number; // 1 - display, 2 - edit, 3 - new
    inlineForm: boolean;
    inlineFormLabelWidth: string;
    onValueChanged (newValue: any): void;
    onEditButtonClick (fieldInternalName: string): void;
    onSave (): any;
    onCancel (fieldName: string): void;
    onAttachmentsChange (newAttachments: any[], oldAttachments: any[]): void;
}