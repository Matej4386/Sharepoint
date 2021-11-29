import * as React from 'react';
import { IMFormsProps } from './IMFormsProps';
import { IMFormsState } from './IMFormsState';
import styles from './MForm.module.scss';
import { List, sp, AttachmentFileAddResult } from '@pnp/sp';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { ScrollablePane } from 'office-ui-fabric-react/lib/ScrollablePane';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Spinner, SpinnerSize} from 'office-ui-fabric-react/lib/Spinner';
import { Overlay } from 'office-ui-fabric-react/lib/Overlay';
import { FormHelpers } from './utilities/FormHelpers';
import { IFormHelpers } from './utilities/IFormHelpers';
import { Fields } from './Field/Fields';
import { IFieldSchema, IFormSchema } from './utilities/RenderListData';
import * as strings from 'MFormWebPartStrings';

export default class MForms extends React.Component < IMFormsProps, IMFormsState > {
    private formHelpers: IFormHelpers;
    constructor(props: IMFormsProps) {
        super(props);
        // Initialize pnp sp
        sp.setup({
          spfxContext: this.props.context
        });
        // Initialize state
        this.state = {
            fieldSchema: [],
            originalFieldSchema: [],
            contenTypeId: '',
            item: null,
            originalItem: null,
            isLoading: true,
            isSaving: false,
            editableFields: [],
            fieldErrors: {},
            requiredFieldEmpty: false,
            spListInfoServerRelativeUrl: '',
            newAttachments: [],
            deleteAttachments: [],
            errors: [],
            notifications: []
        };
        this.formHelpers = new FormHelpers(props.context.spHttpClient);
    }
    public componentDidMount (): void {
        if (this.props.debug) {
            if (this.props.debug === true) {
                console.log ('Form: componentDidMount');
                console.log ('props:');
                console.log (this.props);
                console.log ('state:');
                console.log (this.state);
            }
        }
        this.getInformations();
    }
    public componentDidUpdate (prevProps: IMFormsProps): void {
        if (this.props.debug) {
            if (this.props.debug === true) {
                console.log ('Form: componentDidUpdate');
                console.log ('props:');
                console.log (this.props);
                console.log ('prevProps:');
                console.log (prevProps);
                console.log ('state:');
                console.log (this.state);
            }
        }
        if ((prevProps.listId !== this.props.listId) || (prevProps.listId !== this.props.listId) || (prevProps.formRenderMethod !== this.props.formRenderMethod)) {
            this.getInformations();
        }
    }
    public render(): React.ReactElement<{}> {
        return  <div className={styles.fullwidthContainer}>
                    <div className={styles.container} style={{height: this.props.height ? this.props.height : '650px'}}>
                        {this.renderWaiting()}
                        <ScrollablePane>
                            <div className={styles.containerChild}>
                                {this.renderFields()}
                            </div>
                        </ScrollablePane>
                    </div>
                    {this.renderNotifications()}
                    {this.renderErrors()}
                    {this.renderFooter()}
                </div>;
    }
    private renderFields = (): JSX.Element => {
        if (!this.state.isLoading) {
            if (this.state.fieldSchema && (this.state.fieldSchema.length > 0)) {
                return this.renderAllFields();
            } else {
                return <div className={css(styles.title, 'ms-font-l')}>{strings.Errors.ErrorNoFields}</div>;
            }
        }
    }
    private renderFooter = (): JSX.Element => {
        let element: JSX.Element = null;
        // render save a cancel only if it is not View Form
        if (this.props.formRenderMethod !== 1) {
            element = <div
                style={{display: 'flex', justifyContent: 'center', paddingBottom: '1rem', paddingTop: '1rem'}}
            >
                <PrimaryButton
                    onClick={ () => this.updateItem() }
                    style={{marginRight: '2rem'}}
                    iconProps={{iconName: 'Save'}}
                >
                    {strings.Save}
                </PrimaryButton>
                <DefaultButton
                    onClick={ () => {
                        if (this.props.onCancelForm) {
                            this.props.onCancelForm();
                        }
                    }}
                    iconProps={{iconName: 'Cancel'}}
                >
                    {strings.Cancel}
                </DefaultButton>
            </div>;
        }
        return element;
    }
    private renderAllFields = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.state.item) {
            element = <div>
                    {this.state.fieldSchema.map((field) => {
                        let renderMethod: number = this.props.formRenderMethod;
                        let isSaving: boolean = false;
                        if (this.state.editableFields.filter((fld) => fld === field.InternalName).length > 0) {
                            renderMethod = 2; // edit
                            if (this.state.isSaving === true) {
                                isSaving = true;
                            }
                        }
                        return  <div style={{marginBottom: this.props.inlineForm ? '0.5rem' : '0rem'}}>
                                    <Fields
                                        debug = {this.props.debug ? this.props.debug === true ? true : false : false}
                                        context = {this.props.context}
                                        field = {field}
                                        value = {this.state.item[field.InternalName]}
                                        originalValue = {this.state.originalItem[field.InternalName]}
                                        errorMessage = {this.state.fieldErrors[field.InternalName]}
                                        onValueChanged = {(newValue: any) => this.onValueChanged(newValue, field.InternalName)}
                                        isSaving = {isSaving}
                                        inlineForm = {this.props.inlineForm}
                                        inlineFormLabelWidth = {this.props.inlineFormLabelWidth}
                                        onSave = {() => this.updateItem()}
                                        onCancel = {(fieldName: string) => this.onEditCancel(fieldName)}
                                        onEditButtonClick = {(fieldInternalName: string) => this.onEditFieldClick(fieldInternalName)}
                                        formRenderMethod = { renderMethod }
                                        baseRenderMethod = {this.props.formRenderMethod}
                                        onAttachmentsChange = {this.onAttachmentsChange}
                                        isEditableInViewForm={this.props.isEditableInViewForm}
                                    />
                                </div>;
                    })}
                    </div>;
        }
        return element;
    }
    private onAttachmentsChange = (newAttachments: any[], deleteAttachments: any[]): void => {
        let newSAttachments: any[] = [...this.state.newAttachments];
        let deleteSAttachments: any[] = [...this.state.deleteAttachments];
        let change: boolean = false;
        if (!isEqual(this.state.newAttachments, newAttachments)) {
            newSAttachments = [...newAttachments];
            change = true;
        }
        if (!isEqual(this.state.deleteAttachments, deleteAttachments)) {
            deleteSAttachments = [...deleteAttachments];
            change = true;
        }
        if (change === true) {
            this.setState({
                ...this.state,
                newAttachments: newSAttachments,
                deleteAttachments: deleteSAttachments
            });
        }
    }
    private onEditFieldClick = async (fieldInternalName: string): Promise<void> => {
        if (this.props.isEditableInViewForm) {
            if (this.state.editableFields.length > 0) {
                await this.updateItem();
            } else {
                this.setState({
                    ...this.state,
                    editableFields: [fieldInternalName]
                });
            }
        }
    }
    private onEditCancel = (fieldName: string): void => {
        let fieldErrors: { [fieldName: string]: string } = {...this.state.fieldErrors};
        fieldErrors = {
            ...fieldErrors,
            [fieldName]: ''
        };
        const newItem: any = {
            ...this.state.item,
            [fieldName]: this.state.originalItem[fieldName]
        };
        this.setState({
           item: newItem,
           fieldErrors: {...fieldErrors},
           editableFields: []
        });
    }
    private onValueChanged = (newValue: any, fieldInternalName: string): void => {
        let fieldErrors: { [fieldName: string]: string } = {...this.state.fieldErrors};
        // check requiered
        if ((this.state.fieldSchema.filter((item) => item.InternalName === fieldInternalName)[0].Required) && !newValue) {
            fieldErrors = {
                ...fieldErrors,
                [fieldInternalName]: strings.Errors.RequiredValueMessage
            };
        } else {
            fieldErrors = {
                ...fieldErrors,
                [fieldInternalName]: ''
            };
        }
        if (this.props.onValueChangedHook) {
            const newState: IMFormsState = this.props.onValueChangedHook(newValue, fieldInternalName, fieldErrors, this.state.fieldSchema, this.state.originalFieldSchema)
            this.setState((prevState: IMFormsState) => {
                return {
                    ...prevState,
                    ...newState
                };
            });
        } else {
            this.setState((prevState: IMFormsState) => {
                return {
                    ...prevState,
                    item: { ...prevState.item, [fieldInternalName]: newValue },
                    fieldErrors: {
                        ...fieldErrors
                    }
                };
            });
        }
    }
    private updateItem = async (): Promise<void> => {
        let requiredError: boolean = false;
        let fieldErrors: { [fieldName: string]: string } = {...this.state.fieldErrors};
        // Get form values
        let formValues: {
            FieldName: string;
            FieldValue: any;
            HasException: boolean;
            ErrorMessage: string;
        }[] = this.formHelpers.GetFormValues(this.state.fieldSchema, this.state.item, this.state.originalItem);
        /**
         * Hook onEditFieldsBeforeUpdate
         */
        if (this.props.onEditFieldsBeforeUpdate) {
            const newFormValues:  {
                FieldName: string;
                FieldValue: any;
                HasException: boolean;
                ErrorMessage: string;
            }[] = this.props.onEditFieldsBeforeUpdate(formValues);
            formValues = [...newFormValues];
        }
        if ((formValues.length > 0) || (this.state.newAttachments.length > 0) || (this.state.deleteAttachments.length > 0)) {
            // check required
            for (let i: number = 0; i < this.state.fieldSchema.length; i++) {
                if ((this.state.fieldSchema[i].Required) && (!this.state.item[this.state.fieldSchema[i].InternalName])) {
                    requiredError = true;
                    fieldErrors = {
                        ...fieldErrors,
                        [this.state.fieldSchema[i].InternalName]: strings.Errors.RequiredValueMessage
                    };
                }
            }
            let customError: boolean = false;
            let errorObj: {
                isError: boolean, // If update should continue
                errorText: string; // error text to render in Info
                fieldErrors: {[fieldName: string]: string; } // error text under the field
            } = undefined;
            /**
             * Hook onCustomCheckBeforeUpdate
             */
            if (this.props.onCustomCheckBeforeUpdate) {
                errorObj = this.props.onCustomCheckBeforeUpdate(this.state.item, this.state.fieldSchema, fieldErrors);
                if (errorObj.isError === true) {
                    customError = true;
                }
            }
            if (customError === true) {
                this.setState({
                    ...this.state,
                    fieldErrors: errorObj.fieldErrors,
                    errors: [...this.state.errors, errorObj.errorText]
                });
                return;
            }
            if (requiredError === true) {
                this.setState({
                    ...this.state,
                    fieldErrors: fieldErrors,
                    errors: [...this.state.errors, strings.Errors.FieldsErrorOnSaving],
                    requiredFieldEmpty: requiredError
                });
                return;
            }
            this.setState({
                ...this.state,
                isSaving: true
            });
            try {
                let itemRestResult: any = null;
                /**
                 * NEW FORM
                 */
                if (this.props.formRenderMethod === 3) {
                    // add content type ID for new Item
                    formValues.push({
                        ErrorMessage: null,
                        FieldName: 'ContentTypeId',
                        FieldValue: this.state.contenTypeId,
                        HasException: false
                    });
                    itemRestResult = await this.formHelpers.createItem (this.props.context.pageContext.web.absoluteUrl, this.state.spListInfoServerRelativeUrl, formValues);
                } else {
                    /**
                     * EDIT FORM, or UPDATE IN DISPLAY FORM
                     */
                     itemRestResult = await this.formHelpers.updateItem (this.props.context.pageContext.web.absoluteUrl, this.state.spListInfoServerRelativeUrl, formValues, this.state.originalItem.ID);
                }
                const newState: IMFormsState = { ...this.state, fieldErrors: {} };
                let hadErrors: boolean = false;
                itemRestResult.filter((fieldVal) => fieldVal.HasException).forEach((element) => {
                    newState.fieldErrors[element.FieldName] = element.ErrorMessage;
                    hadErrors = true;
                });
                let itemId: number = this.props.listItemId ? this.props.listItemId : 0;
                if (hadErrors) {
                   newState.errors = [...newState.errors, strings.Errors.FieldsErrorOnSaving];
                } else {
                    itemRestResult.reduce(
                        (val: any, merged: any) => {
                            merged[val.FieldName] = merged[val.FieldValue]; return merged;
                        },
                        newState.item
                    );
                    /**
                     * New form NEW ID
                     */
                     if (this.props.formRenderMethod === 3) {
                        const idField: any[] = itemRestResult.filter((fieldVal) => fieldVal.FieldName === 'Id');
                        if (idField.length > 0) {
                            itemId = Number(idField[0].FieldValue);
                        }
                    }
                    /**
                     * For Note AppendOnly fields why must load data
                     */
                    const multiLineAppendFields: IFieldSchema[] = this.state.fieldSchema.filter((field) => field.Type === 'Note').filter((note) => note.AppendOnly === true);
                    if (multiLineAppendFields.length > 0) {
                        for (let w: number = 0; w < itemRestResult.length; w++) {
                            if (multiLineAppendFields.filter((multiField) => multiField.InternalName === itemRestResult[w].FieldName).length > 0) {
                                const itemRet: any[] = await sp.web.lists.getById(this.props.listId).items.getById(itemId).versions
                                    .select(itemRestResult[w].FieldName, 'Created', 'Editor')
                                    .filter(itemRestResult[w].FieldName + ' ne null')
                                    .get();
                                if (itemRet) {
                                    if (itemRet.length > 0) {
                                        newState.item[itemRestResult[w].FieldName] = [...itemRet];
                                    } else {
                                        newState.item[itemRestResult[w].FieldName] = [];
                                    }
                                } else {
                                    newState.item[itemRestResult[w].FieldName] = [];
                                }
                            }
                        }
                    }
                    if (this.state.deleteAttachments.length > 0) {
                        for (let i: number = 0; i < this.state.deleteAttachments.length; i++) {
                            const attachmentDelete: any = await sp.web.lists.getById(this.props.listId).items.getById(itemId).attachmentFiles.getByName(this.state.deleteAttachments[i]).recycle();
                            if (this.props.debug) {
                                if (this.props.debug === true) {
                                    console.log ('Form (updateItem): attachmentDelete');
                                    console.log (attachmentDelete);
                                }
                            }
                        }
                        const newAtt: any[] = newState.item.Attachments.Attachments.filter((at: any) => {
                            return !(this.state.deleteAttachments.indexOf(at.FileName) > -1);
                        });
                        newState.item.Attachments.Attachments = newAtt;
                        newState.deleteAttachments = [];
                    }
                    if (this.state.newAttachments.length > 0) {
                        for (let i: number = 0; i < this.state.newAttachments.length; i++) {
                            let fileName: string = this.state.newAttachments[i].file.name;
                            const fileServerRelativeUrl: string = `${this.state.spListInfoServerRelativeUrl}/Attachments/${itemId}/${fileName}`;
                            /**
                             * Check if attachment file already exists in Item - but not for NEW form
                             */
                            if (this.props.formRenderMethod !== 3) {
                                const exists: any = await sp.web
                                    .getFileByServerRelativeUrl(fileServerRelativeUrl)
                                    .select('Exists').get()
                                    .then((d) => d.Exists)
                                    .catch(() => false);
                                if (exists === true) {
                                    fileName = 'Dupl_' + this.state.newAttachments[i].file.name;
                                    this.renderInfo (true, strings.Errors.ErrorDuplicateAttachment);
                                }
                            }
                            const addAttachment: AttachmentFileAddResult = await sp.web.lists.getById(this.props.listId).items.getById(itemId).attachmentFiles.add(fileName, this.state.newAttachments[i].file);
                            if (this.props.debug) {
                                if (this.props.debug === true) {
                                    console.log ('Form (updateItem): addAttachment');
                                    console.log (addAttachment);
                                }
                            }
                            if ((newState.item.Attachments === null) || (newState.item.Attachments === '') || (newState.item.Attachments === undefined)) {
                                let urlPrefix: string = '';
                                if (newState.item.EncodedAbsUrl) {
                                    const split: string[] = newState.item.EncodedAbsUrl.split('/');
                                    urlPrefix = split.slice(0, split.length - 1).join('/') + '/Attachments/' + itemId.toString() + '/';
                                } else {
                                    const encodedAbsUrl: any = await sp.web.lists.getById(this.props.listId).items.getById(itemId).select('EncodedAbsUrl').get();
                                    const split: string[] = encodedAbsUrl.EncodedAbsUrl.split('/');
                                    urlPrefix = split.slice(0, split.length - 1).join('/') + '/Attachments/' + itemId.toString() + '/';
                                }
                                newState.item.Attachments = {
                                    Attachments: [{
                                        AttachmentId: null,
                                        FileName: addAttachment.data.FileName,
                                        FileTypeProgId: null,
                                        RedirectUrl: null
                                    }],
                                    UrlPrefix: urlPrefix
                                };
                            } else {
                                newState.item.Attachments.Attachments.push({
                                    AttachmentId: null,
                                    FileName: addAttachment.data.FileName,
                                    FileTypeProgId: null,
                                    RedirectUrl: null
                                });
                            }
                        }
                        newState.newAttachments = [];
                    }
                    newState.originalItem = { ...newState.item };
                    newState.editableFields = [];
                    newState.notifications = [...this.state.notifications, strings.ItemSavedSuccessfully];
                }
                newState.isSaving = false;
                this.setState(newState, () => {
                    if (this.props.onSuccesSave) {
                        this.props.onSuccesSave(this.state.item, itemId);
                    }
                });
            } catch (error) {
                const errorText: string = `(saveItem) -> ${strings.Errors.ErrorOnSavingListItem}: ${error}`;
                this.renderInfo (true, errorText);
                this.setState({ ...this.state, isSaving: false}, () => {
                    if (this.props.onErrorSave) {
                        this.props.onErrorSave (this.state.item, error);
                    }
                });
            }
        } else {
            this.setState({
                editableFields: []
            });
        }
    }
    /**
     * Get informations about form, item, ...
     * Load form schema
     * Load item
     * Load append fields
     * Hooks
     */
    private getInformations = async (): Promise<void> => {
        let contentTypeId: string = this.props.contentTypeId;
        try {
            const spList: List = sp.web.lists.getById(this.props.listId);
            const spListInfo: any = await spList.rootFolder.get();
            let item: any = undefined;

            const myObjectForm: IFormSchema = {
                ContentTypeId: '',
                formFields: []
            };
            const renderListDataParams: any = {
                ViewXml: '<View><ViewFields><FieldRef Name="ID"/></ViewFields></View>',
                RenderOptions: 64
            };
            const response: any = await spList.renderListDataAsStream(renderListDataParams);
            /**
             * Check if attachments are enabled in the lits
             */
            const attachmentEnabled: boolean = ('EnableAttachments' in response) ?
                    response.EnableAttachments === 'true' ?
                        true
                        :
                        false
                    :
                    false;
            const form: any =  response.ClientForms.New;
            const contentTypes: any = response.ContentTypeIdToNameMap;
            let contentypeName: string = null;
            /**
             * Select Content type name for selected Content type from form settings
             */
            if (Object.keys(contentTypes).length) {
                Object.keys(contentTypes).forEach(key => {
                    if (this.props.contentTypeId) {
                        if (this.props.contentTypeId === key) {
                            myObjectForm.ContentTypeId = key;
                            contentypeName = contentTypes[key];
                        }
                    } else {
                        myObjectForm.ContentTypeId = key;
                    }
                });
            }
            /**
             * Select form fields for selected content type
             */
            if (Object.keys(form).length) {
                Object.keys(form).forEach(key => {
                    if (contentypeName !== null) {
                        if (key === contentypeName) {
                            myObjectForm.formFields = form[key];
                        }
                    } else {
                        myObjectForm.formFields = form[key];
                    }
                });
            }

            // if Attachments are disabled filter Attachments field from schema
            if (attachmentEnabled === false) {
                myObjectForm.formFields = myObjectForm.formFields.filter((field) => field.Type !== 'Attachments');
            }
            /**
             * Load default content type ID
             */
            if (contentTypeId === undefined || contentTypeId === '') {
                const defaultContentType: any[] = await spList.contentTypes.select('Id', 'Name').get();
                contentTypeId = defaultContentType[0]['Id'].StringValue;
                myObjectForm.ContentTypeId = contentTypeId;
            }
            /**
             * Load list Item
             */
            if (this.props.listItemId !== undefined && this.props.listItemId !== null && this.props.listItemId !== 0 && (this.props.formRenderMethod !== 3)) {
                item = await this.formHelpers.getDataForForm(this.props.context.pageContext.web.absoluteUrl, spListInfo.ServerRelativeUrl, this.props.listItemId);
                if (item !== null) {
                    // For multiline append fields - Sharepoint onpremise bug - does not return append versions
                    const multiLineAppendFields: IFieldSchema[] = myObjectForm.formFields.filter((field) => field.Type === 'Note').filter((note) => note.AppendOnly === true);
                    if (multiLineAppendFields.length > 0) {
                        for (let i: number = 0; i < multiLineAppendFields.length; i++) {
                            const itemRet: any[] = await spList.items.getById(this.props.listItemId).versions
                                .select(multiLineAppendFields[i].InternalName, 'Created', 'Editor')
                                .filter(multiLineAppendFields[i].InternalName + ' ne null')
                                .get();
                            if (itemRet) {
                                if (itemRet.length > 0) {
                                    item[multiLineAppendFields[i].InternalName] = [...itemRet];
                                }
                            }
                        }
                    }
                    // for URL
                    const noteFields: IFieldSchema[] = myObjectForm.formFields.filter((field) => field.Type === 'URL');
                    if (noteFields.length > 0) {
                        for (let i: number = 0; i < noteFields.length; i++) {
                            item[noteFields[i].InternalName] = {
                                URL: item[noteFields[i].InternalName],
                                Desc: item[noteFields[i].InternalName + '.desc']
                            };
                        }
                    }
                }
            } else if (this.props.formRenderMethod === 3) {
                /**
                 * Load default values for new form
                 */
                let data: any = {};
                for (let i: number = 0; i < myObjectForm.formFields.length; i++) {
                    if (myObjectForm.formFields[i].DefaultValue) {
                    data = {
                        ...data,
                        [myObjectForm.formFields[i].InternalName]: myObjectForm.formFields[i].DefaultValue
                    };
                    }
                }
                item = {...data};
                // for URL
                const noteFields: IFieldSchema[] = myObjectForm.formFields.filter((field) => field.Type === 'URL');
                if (noteFields.length > 0) {
                    for (let i: number = 0; i < noteFields.length; i++) {
                        item[noteFields[i].InternalName] = {
                            URL: '',
                            Desc: ''
                        };
                    }
                }
            }
            if (this.props.onLoadChangeItem) {
                const newitem: any = this.props.onLoadChangeItem(item);
                item = {...newitem};
            }
            if (this.props.onLoadChangeFieldSchema) {
                const newFormfields: IFieldSchema[] = this.props.onLoadChangeFieldSchema(myObjectForm.formFields);
                myObjectForm.formFields = [...newFormfields];
            }
            this.setState({
                isLoading: false,
                item: item,
                originalItem: item,
                fieldSchema: myObjectForm.formFields,
                originalFieldSchema: myObjectForm.formFields,
                contenTypeId: myObjectForm.ContentTypeId,
                spListInfoServerRelativeUrl: spListInfo.ServerRelativeUrl,
                editableFields: [],
                fieldErrors: {},
                requiredFieldEmpty: false,
                newAttachments: [],
                deleteAttachments: [],
                errors: [],
                notifications: []
            });
        } catch (error) {
            const errorText: string = `(Loading form) -> ${strings.Errors.ErrorOnLoadingApp}: ${error}`;
            this.renderInfo (true, errorText);
            this.setState({
                ...this.state,
                isLoading: false
            });
        }
    }
    private renderWaiting = (): JSX.Element => {
        return ((this.state.isLoading === true) || ((this.state.isSaving === true) && (this.props.formRenderMethod !== 1))) ?
            <div>
                <Overlay isDarkThemed={false} className={styles.overlay}>
                    <Spinner size={SpinnerSize.medium} />
                </Overlay>
            </div>
            :
            null;
    }
    private renderInfo = (error: boolean, message: string): void => {
        if (error === true) {
          this.setState({
            ...this.state,
            errors: [...this.state.errors, message]
          });
        } else {
            this.setState({
                ...this.state,
                notifications: [...this.state.notifications, message]
            });
        }
    }
    private renderNotifications = (): JSX.Element => {
        if (this.state.notifications.length === 0) {
          return null;
        }
        setTimeout(() => { this.setState({ ...this.state, notifications: [] }); }, 4000);
        return <div className={styles.Footercontainer}>
          {
            this.state.notifications.map((item) =>
              <MessageBar messageBarType={MessageBarType.success}>{item}</MessageBar>
            )
          }
        </div>;
      }
    private renderErrors = (): JSX.Element => {
        return this.state.errors.length > 0 ?
          <div className={styles.Footercontainer}>
            <div>
              {
                this.state.errors.map((item, idx) =>
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={true}
                    onDismiss={() => this.clearError(idx)}
                  >
                    {item}
                  </MessageBar>
                )
              }
            </div>
          </div>
          :
          null;
    }
    private clearError = (idx: number): void => {
        const newErrors: string[] = this.state.errors.splice(idx + 1, 1);
        this.setState({
          ...this.state,
          errors: newErrors
        });
    }
}