import * as React from 'react';
import ardStyles from '../MForm.module.scss';
import * as stylesImport from 'office-ui-fabric-react/lib/components/TextField/TextField.types';
const styles: any = stylesImport;
import * as moment from 'moment';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { IFieldsProps } from './IFieldsProps';
import { IFieldsState } from './IFieldsState';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { css, DelayedRender } from 'office-ui-fabric-react/lib/Utilities';
import ReactHtmlParser from 'react-html-parser';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { FieldUserRenderer } from '@pnp/spfx-controls-react/lib/FieldUserRenderer';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { DateTimePicker, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { FieldDateRenderer } from '@pnp/spfx-controls-react/lib/FieldDateRenderer';
import { Overlay } from 'office-ui-fabric-react/lib/Overlay';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Locales } from '../utilities/Locales';
import * as strings from 'MFormWebPartStrings';
import { FieldTextRenderer } from '@pnp/spfx-controls-react/lib/FieldTextRenderer';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IconButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { TaxonomyPicker, IPickerTerms, IPickerTerm } from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
import utilities from '../utilities/utilities';

export class Fields extends React.Component<IFieldsProps, IFieldsState> {
    private newAttachmentsRefs: any = null;
    private _utilities: utilities;
    constructor(props: IFieldsProps) {
        super(props);
        this._utilities = new utilities();
        // Initialize state
        this.state = {
            showEdit: false,
            newAttachments: [],
            deleteAttachments: []
        };
    }
    public componentWillReceiveProps (): void {
        if (this.props.debug) {
            if (this.props.debug === true) {
                console.log ('Field: componentWillReceiveProps');
                console.log ('props:');
                console.log (this.props);
                console.log ('state:');
                console.log (this.state);
            }
        }
        this.newAttachmentsRefs = null;
    }
    public componentDidMount(): void {
        if (this.props.debug) {
            if (this.props.debug === true) {
                console.log ('Field: componentDidMount');
                console.log ('props:');
                console.log (this.props);
                console.log ('state:');
                console.log (this.state);
            }
        }
    }
    public componentDidUpdate(prevProps: IFieldsProps, prevState: IFieldsState ): void {
        if (prevProps.formRenderMethod !== this.props.formRenderMethod) {
            if (this.props.formRenderMethod !== 1) {
                this.setState({
                    showEdit: false
                });
            }
            this.setState({
                newAttachments: [],
                deleteAttachments: []
            });
        }
        const newAttachments: {file: any}[] = this.state.newAttachments.filter(
            (attachment) => attachment.file !== null
        );

        if ((newAttachments.length > 0) || (!isEqual(this.state.deleteAttachments, prevState.deleteAttachments)) || (!isEqual(this.state.newAttachments, prevState.newAttachments))) {
            this.props.onAttachmentsChange(newAttachments, this.state.deleteAttachments);
        }
    }
    public render(): React.ReactElement<{}> {
        let element: JSX.Element = undefined;
        if (this.props.formRenderMethod === 1) {// display
            element =   <div className={css(ardStyles.FormFieldfieldGroup, styles.formFieldClassName)} style={{display: 'inline-flex', width: '100%'}}>
                            <div className={this.props.inlineForm ? ardStyles.inlineForms : styles.wrapper} style={{width: this.props.isEditableInViewForm ? '90%' : '100%'}}>
                                <div style={{display: 'inline-flex', width: this.props.inlineForm ? this.props.inlineFormLabelWidth : ''}}>
                                    {this.renderIcon()}
                                    <Label className={css(ardStyles.label, { ['is-required']: this.props.field.Required })}>{this.props.field.Title}</Label>
                                </div>
                                <div
                                    onMouseEnter={() => this.setState({showEdit: true})}
                                    onMouseOut={() => this.setState({showEdit: false})}
                                    className={ardStyles.controlContainerDisplay}
                                    onClick={() => (this.props.field.Type === 'Attachments' || this.props.field.Type === 'URL') ?
                                        null
                                        :
                                        this.props.isEditableInViewForm ? this.props.onEditButtonClick(this.props.field.InternalName) : null
                                    }
                                >
                                    {this.renderField()}
                                </div>
                            </div>
                            { (((this.state.showEdit) || (this.props.field.Type === 'Attachments') || (this.props.field.Type === 'URL')) && (this.props.isEditableInViewForm)) &&
                                <div className={ardStyles.IconEdit}>
                                    <IconButton onClick={() => this.props.onEditButtonClick(this.props.field.InternalName)} iconProps={{iconName: 'Edit'}} style={{paddingTop: '5px', paddingBottom: '5px', paddingLeft: '1rem'}}/>
                                </div>
                            }
                        </div>;
        } else { // edit , new
            element =   <div className={ardStyles.FormFieldfieldGroup}>
                            <div className={css(styles.formFieldClassName)} style={{display: 'inline-flex', width: '100%'}}>
                                <div className={this.props.inlineForm ? ardStyles.inlineForms : styles.wrapper} style={{width: this.props.baseRenderMethod !== 1 ? '100%' : '90%'}}>
                                    <div style={{display: 'inline-flex', width: this.props.inlineForm ? this.props.inlineFormLabelWidth : ''}}>
                                        {this.renderIcon()}
                                        <Label className={css(ardStyles.label, { ['is-required']: this.props.field.Required })}>{this.props.field.Title}</Label>
                                    </div>
                                    <div className={ardStyles.controlContainerDisplay}>
                                        {this.renderWaiting()}
                                        {this.renderField()}
                                        {this.props.baseRenderMethod === 1 && // only if it was originaly view form
                                            <div className={ardStyles.IconEdit}>
                                                <IconButton
                                                    iconProps={ { iconName: 'CheckMark' } }
                                                    title={strings.Save}
                                                    style={{fontWeight: 'bold'}}
                                                    ariaLabel={strings.Save}
                                                    onClick={() => this.props.onSave()}
                                                    disabled={this.props.isSaving}
                                                />
                                                <IconButton
                                                    iconProps={ { iconName: 'Cancel' } }
                                                    title={strings.Cancel}
                                                    style={{fontWeight: 'bold'}}
                                                    ariaLabel={strings.Cancel}
                                                    onClick={() => {
                                                            this.newAttachmentsRefs = null;
                                                            this.props.onCancel(this.props.field.InternalName);
                                                        }
                                                    }
                                                    disabled={this.props.isSaving}
                                                />
                                            </div>
                                        }
                                        <div>
                                        {
                                            this.props.field.AppendOnly && this.props.formRenderMethod === 2 && this.renderAppend(this.props)
                                        }
                                        </div>
                                    </div>
                                </div>
                            </div>
                            {this.props.errorMessage &&
                            <div className={ardStyles.controlContainerDisplay} style={{paddingTop: '0px'}}>
                                <div aria-live='assertive'>
                                    <DelayedRender>
                                        <p className={ardStyles.myErrorMessage}>
                                            {Icon({ iconName: 'Error', className: styles.errorIcon })}
                                            <span className={styles.errorText} data-automation-id='error-message'>{this.props.errorMessage}</span>
                                        </p>
                                    </DelayedRender>
                                </div>
                            </div>
                            }
                        </div>;
        }
        return element;
    }
    private renderIcon = () => {
        let field: JSX.Element = undefined;
        switch (this.props.field.Type) {
            case 'Text':
                field = <Icon className={ardStyles.fieldIcon} iconName={'TextField'} />;
                break;
            case 'Note':
                field = <Icon className={ardStyles.fieldIcon} iconName={'AlignLeft'} />;
                break;
            case 'User':
            case 'UserMulti':
                field = <Icon className={ardStyles.fieldIcon} iconName={'Contact'} />;
                break;
            case 'Boolean':
                field = <Icon className={ardStyles.fieldIcon} iconName={'CheckboxComposite'} />;
                break;
            case 'DateTime':
                field = <Icon className={ardStyles.fieldIcon} iconName={'Calendar'} />;
                break;
            case 'Choice':
                field = <Icon className={ardStyles.fieldIcon} iconName={'CheckMark'} />;
                break;
            case 'MultiChoice':
                field = <Icon className={ardStyles.fieldIcon} iconName={'MultiSelect'} />;
                break;
            case 'Number':
                field = <Icon className={ardStyles.fieldIcon} iconName={'NumberField'} />;
                break;
            case 'Currency':
                field = <Icon className={ardStyles.fieldIcon} iconName={'Money'} />;
                break;
            case 'Attachments':
                field = <Icon className={ardStyles.fieldIcon} iconName={'Attach'} />;
                break;
            case 'Lookup':
                switch (this.props.field.FieldType) {
                    case 'TaxonomyFieldType':
                    case 'TaxonomyFieldTypeMulti':
                        field = <Icon className={ardStyles.fieldIcon} iconName={'BulletedList'} />;
                        break;
                    case 'Lookup':
                    case 'LookupMulti':
                        field = <Icon className={ardStyles.fieldIcon} iconName={'Switch'} />;
                        break;
                    default:
                        field = null;
                        break;
                }
                break;
            case 'URL':
                field = <Icon className={ardStyles.fieldIcon} iconName={'Link'} />;
                break;
            default:
                field = null;
                break;
        }
        return field;
    }
    private renderField = () => {
        let field: JSX.Element = undefined;
        switch (this.props.field.Type) {
            case 'Text':
                field = this.renderText();
                break;
            case 'Note':
                field = this.renderNote();
                break;
            case 'User':
            case 'UserMulti':
                field = this.renderUser();
                break;
            case 'Boolean':
                field = this.renderBoolean();
                break;
            case 'DateTime':
                field = this.renderDate();
                break;
            case 'Choice':
            case 'MultiChoice':
                field = this.renderChoice();
                break;
            case 'Number':
            case 'Currency':
                field = this.renderNumber();
                break;
            case 'Attachments':
                field = this.renderAttachments();
                break;
            case 'Lookup':
                switch (this.props.field.FieldType) {
                    case 'TaxonomyFieldType':
                    case 'TaxonomyFieldTypeMulti':
                        field = this.renderTaxonomy (this.props.field.FieldType);
                        break;
                    case 'Lookup':
                    case 'LookupMulti':
                        field = this.renderLookup();
                        break;
                    default:
                        field = <div>{this.props.field.Type}</div>;
                        break;
                }
                break;
            case 'URL':
                field = this.renderHypertext();
                break;
            default:
                field = <div>Not supported field: {this.props.field.Type}</div>;
                break;
        }
        return field;
    }
    private renderLookup = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            if ((this.props.value) && (this.props.value.length > 0)) {

                const baseUrl: string = `${this.props.field.BaseDisplayFormUrl}&ListId={${this.props.field.LookupListId}}`;
                let values: any = this.props.value;
                if (!Array.isArray(this.props.value)) {
                    if (this.props.field.FieldType !== 'LookupMulti') {
                        const splitArray: string[] = this.props.value.split(';#');
                        values = splitArray.filter((item, idx) => (idx % 2 === 0))
                            .map((comp, idx) => ({ lookupId: Number(comp), lookupValue: (splitArray.length > idx + 1) ? splitArray[idx + 1] : '' }));
                    } else {
                        const splitArray: string[] = this.props.value.split(';#_');
                        let val: any;
                        values = [];
                        for (let i: number = 0; i < splitArray.length; i++) {
                            const split: string[] = splitArray[i].split(';#');
                            if (split.length === 2) {
                                val = {
                                    lookupId: Number(split[0]),
                                    lookupValue: split[1]
                                };
                                values.push(val);
                            }
                        }
                    }
                }
                element = <div>
                    {
                        values.map((val) => <div><Link target={'_blank'} href={`${baseUrl}&ID=${val.lookupId}`}>{val.lookupValue}</Link></div>)
                    }
                </div>;
            } else {
                element = <div></div>;
            }
        } else {

            let options: {
                key: any;
                text: any;
            }[] = this.props.field.Choices.map((option) => ({ key: option.LookupId, text: option.LookupValue }));
            if (this.props.field.FieldType !== 'LookupMulti') {
                if (!this.props.field.Required) { options = [{ key: 0, text: '' }].concat(options); }
                const value: any = this.props.value ?
                    Array.isArray(this.props.value) ?
                        this.props.value[0].lookupId
                        :
                        Number(this.props.value.split(';#')[0])
                    :
                    0;
                element = <Dropdown
                    className={styles.dropDownFormField}
                    options={options}
                    selectedKey={value}
                    onChanged={(item) => this.props.onValueChanged(`${item.key};#${item.text}`)}
                />;
            } else {
                let values: any[] = [];
                if (this.props.value) {
                    if (Array.isArray(this.props.value)) {
                        values = this.props.value.map((val) => ({ key: Number(val.lookupId), text: val.lookupValue}));
                    } else {
                        const splitArray: string[] = this.props.value.split(';#_');
                        let val: any;
                        for (let i: number = 0; i < splitArray.length; i++) {
                            const split: string[] = splitArray[i].split(';#');
                            if (split.length === 2) {
                                val = {
                                    key: Number(split[0]),
                                    text: split[1]
                                };
                                values.push(val);
                            }
                        }
                    }
                }
                element = <Dropdown
                    className={styles.dropDownFormField}
                    options={options}
                    selectedKeys={values.map((val) => val.key)}
                    multiSelect
                    onChanged={(item) => this.props.onValueChanged(getUpdatedValue(values, item))}
                />;
            }
        }
        return element;

        function getUpdatedValue(oldValues: Array<{ key: number, text: string }>, changedItem: IDropdownOption): string {
            let newValues: Array<{ key: number, text: string }>;
            if (changedItem.selected) {
                newValues = [...oldValues, { key: Number(changedItem.key), text: changedItem.text }];
            } else {
                newValues = oldValues.filter((item) => item.key !== changedItem.key);
            }
            return newValues.reduce((valStr, item) => valStr + `${item.key};#${item.text};#_`, '');
        }
    }
    private renderHypertext = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            const value: any = this.props.value ? this.props.value : {
                URL: '',
                Desc: ''
            };
            if (this.props.field.DisplayFormat === 0) { // URL
                element = <Link
                    target={'_blank'}
                    href={value.URL}
                >
                    {value.URL !== '' ? value.Desc !== '' ? value.Desc : value.URL : ''}
                </Link>;
            } else { // Image
                if (value.URL !== '') {
                    element = <div>
                        <Image
                            src={value.URL !== '' ? value.URL : ''}
                            height={100}
                            width={100}
                            imageFit={ImageFit.contain}
                        />
                        <Link
                            target={'_blank'}
                            href={value.URL}
                        >
                            {value.URL !== '' ? value.Desc !== '' ? value.Desc : value.URL : ''}
                        </Link>
                    </div>;
                } else {
                    element = null;
                }
            }
        } else {
            element = <div>
           <TextField
                name={this.props.field.InternalName}
                value={this.props.value.URL}
                onChanged={(newValue) => this.onURLChange(newValue, true)}
                placeholder={strings.UrlFormFieldPlaceholder}
                multiline={false}
                style={{marginBottom: '0.3rem'}}
                validateOnFocusIn
                validateOnFocusOut
            />
            <TextField
                name={this.props.field.InternalName + '.desc'}
                value={this.props.value.Desc}
                onChanged={(newValue) => this.onURLChange(newValue, false)}
                placeholder={strings.UrlDescFormFieldPlaceholder}
                multiline={false}
                validateOnFocusIn
                validateOnFocusOut
            />
          </div>;
        }
        return element;
    }
    private onURLChange = (value: string, isUrl: boolean) => {
        let currValue: any = this.props.value || {
            URL: '',
            Desc: ''
        };
        currValue = {
          ...currValue
        };

        if (isUrl) {
          currValue.URL = value;
        } else {
          currValue.Desc = value;
        }
        this.props.onValueChanged(currValue);
      }
    private renderTaxonomy = (fieldType: string): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            let value: string = '';
            if (fieldType === 'TaxonomyFieldType') {
                value = (this.props.value) ?
                    (typeof this.props.value.Label === 'string') ?
                        this.props.value.Label
                        :
                        (typeof this.props.value.name === 'string') ?
                            this.props.value.name
                            :
                            Array.isArray(this.props.value) ?
                                this.props.value.map((v) => v.name).join(', ')
                                :
                                JSON.stringify(this.props.value)
                    :
                    '';
            } else { // TaxonomyFieldTypeMulti
                value = (this.props.value) ?
                    this.props.value.length > 0 ?
                        (typeof this.props.value[0].Label === 'string') ?
                            this.props.value.map((v) => v.Label).join(', ')
                            :
                            (typeof this.props.value[0].name === 'string') ?
                                this.props.value.map((v) => v.name).join(', ')
                                :
                                ''
                        :
                        ''
                    :
                    '';
            }
            element = <FieldTextRenderer text={value} />;
        } else {
            const value: IPickerTerms = [];
            if (this.props.value) {
                if ((typeof this.props.value === 'object') && (!Array.isArray(this.props.value))) {
                    if ('key' in this.props.value) {
                        if (this.props.value.key !== undefined) {
                            value.push(this.props.value);
                        }
                    } else {
                        const iPickerTermObj: IPickerTerm = {
                            key: '',
                            name: '',
                            path: '',
                            termSet: ''
                        };
                        iPickerTermObj.key = this.props.value.TermID;
                        iPickerTermObj.name = this.props.value.Label;
                        iPickerTermObj.termSet = this.props.field.TermSetId;
                        value.push(iPickerTermObj);
                    }
                } else if (Array.isArray(this.props.value)) {
                    for (let i: number = 0; i < this.props.value.length; i++) {
                        if ('key' in this.props.value[i]) {
                            if (this.props.value[i].key !== undefined) {
                                value.push(this.props.value[i]);
                            }
                        } else {
                            const iPickerTermObj: IPickerTerm = {
                                key: '',
                                name: '',
                                path: '',
                                termSet: ''
                            };
                            iPickerTermObj.key = this.props.value[i].TermID;
                            iPickerTermObj.name = this.props.value[i].Label;
                            iPickerTermObj.termSet = this.props.field.TermSetId;
                            value.push(iPickerTermObj);
                        }
                    }
                }
            }
            element = <TaxonomyPicker
                allowMultipleSelections={fieldType === 'TaxonomyFieldType' ? false : true}
                termsetNameOrID={this.props.field.TermSetId}
                panelTitle={this.props.field.Title}
                initialValues={value}
                label=''
                context={this.props.context}
                onChange={(terms: IPickerTerms) => {
                    this.props.onValueChanged(terms);
                }}
                isTermSetSelectable={false}
            />;
        }
        return element;
    }
    private renderText = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            const value: string = (this.props.value) ? ((typeof this.props.value === 'string') ? this.props.value : JSON.stringify(this.props.value)) : '';
            element = <FieldTextRenderer
                        text={value}
                    />;
        } else {
            const value: string = this.props.value ? this.props.value : '';
            element = <TextField
                name={this.props.field.InternalName}
                value={value}
                onChanged={(newValue) => this.props.onValueChanged(newValue)}
                placeholder={strings.TextFormFieldPlaceholder}
                multiline={false}
                validateOnFocusIn
                validateOnFocusOut
            />;
        }
        return element;
    }
    private renderNote = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            if (this.props.field.AppendOnly === true) {
                element = this.renderAppend(this.props);
            } else {
                if (this.props.field.RichText === true) {
                    element = <div>{ReactHtmlParser(this.props.value)}</div>;
                } else {
                    const value: string = (this.props.value) ? ((typeof this.props.value === 'string') ? this.props.value : JSON.stringify(this.props.value)) : '';
                    element = <FieldTextRenderer
                            text={value}
                        />;
                }
            }
        } else {
            if (this.props.field.RichText === true) {
                const value: any = this.props.value ? this.props.value : '';

                element = <div><RichText
                    value={this.props.field.AppendOnly ? '' : value}
                    onChange={(text) => {
                        this.props.onValueChanged(text);
                        return text;
                    }}
                    placeholder={strings.TextFormFieldPlaceholder}
                    className={ardStyles.richTextEditor}
                    isEditMode={true}
                />
                </div>;
            } else {
                let value: string = this.props.value ? this.props.value : '';
                if (this.props.field.AppendOnly) {
                    if (Array.isArray(value)) {
                        value = '';
                    }
                }
                element = <div><TextField
                    name={this.props.field.InternalName}
                    value={value}
                    onChanged={(text) => {
                            this.props.onValueChanged(text);
                    }}
                    placeholder={strings.TextFormFieldPlaceholder}
                    multiline={true}
                    validateOnFocusIn
                    validateOnFocusOut
                />
                </div>;
            }
        }

        return element;
    }
    private renderAppend = (props: IFieldsProps): JSX.Element => {
        let el: JSX.Element = <div></div>;
            if (Array.isArray(props.originalValue)) {
                const locale: string = Locales[props.field.LocaleId];
                moment.locale(locale);
                el = <div>
                    {props.originalValue.map((append) => {
                        let date: string = '';
                        date = (append.Created && moment(append.Created).isValid()) ? moment.utc(append.Created).local().format('L HH:mm') : '';
                        let val: JSX.Element = null;
                        val =   <div style={{margin: '0.3rem', marginBottom: '0.8rem'}}>
                                    <div style={{display: 'inline-flex'}}>
                                        <div style={{fontWeight: 'bold'}}>{append.Editor.LookupValue}</div>
                                        <div style={{marginLeft: '0.3rem'}}>({date}) :</div>
                                    </div>
                                    <div>{ReactHtmlParser(append[props.field.InternalName])}</div>
                                </div>;
                        return val;
                    })}
                </div>;
            }
            return el;
    }
    private renderUser = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            if ((this.props.value) && (this.props.value.length > 0)) {
                const newvalue: any[] = [];
                for (let q: number = 0; q < this.props.value.length; q++) {
                    if (this.props.value[q].hasOwnProperty('secondaryText')) {
                        const obj: any = {
                            id: this.props.value[q].id,
                            email: this.props.value[q].secondaryText,
                            department: '',
                            jobTitle: '',
                            sip: this.props.value[q].secondaryText,
                            title: this.props.value[q].text,
                            value: this.props.value[q].loginName,
                            picture: this.props.value[q].imageUrl
                        };
                        newvalue.push(obj);
                    } else {
                        newvalue.push(this.props.value[q]);
                    }
                }
                element = <FieldUserRenderer users={newvalue} context={this.props.context} />;
            } else {
                element = <div></div>;
            }
        } else {
            const value: string[] = [];
            if (this.props.value) {
                if (this.props.value && (this.props.value.length > 0)) {
                    for (let k: number = 0; k < this.props.value.length; k++) {
                        if (this.props.value[k].hasOwnProperty('email')) {
                            value.push(this.props.value[k].email);
                        } else if (this.props.value[k].hasOwnProperty('secondaryText')) {
                            value.push(this.props.value[k].secondaryText);
                        } else {
                            value.push('');
                        }
                    }
                } else {
                    value.push('');
                }
            } else {  // else we do not have any default selected user
                value.push('');
            }
            element = <PeoplePicker
                context={this.props.context}
                personSelectionLimit={this.props.field.AllowMultipleValues === true ? 30 : 1}
                groupName={''} // Leave this blank in case you want to filter from all users
                showtooltip={false}
                // isRequired={props.fieldSchema.Required}
                // errorMessage={strings.FormFields.RequiredValueMessage}
                defaultSelectedUsers={value}
                disabled={false}
                ensureUser={true}
                placeholder={this.props.field.AllowMultipleValues === true ? strings.UsersFormFieldPlaceholder : strings.UserFormFieldPlaceholder}
                selectedItems={(person: any[]) => {this.props.onValueChanged(person); }}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={500}
            />;
        }
        return element;
    }
    private renderBoolean = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            let value: string = (this.props.value) ? ((typeof this.props.value === 'string') ? this.props.value : JSON.stringify(this.props.value)) : '';
            value = (this.props.value === '1' || this.props.value === 'true' || this.props.value === 'Yes' || this.props.value === 'Áno') ? strings.ToggleOnText : strings.ToggleOffText;
            element = <FieldTextRenderer
                        text={value}
                    />;
        } else {
            element = <Toggle
                checked={this.props.value === '1' || this.props.value === 'true' || this.props.value === 'Yes' || this.props.value === 'Áno'}
                onAriaLabel={strings.ToggleOnAriaLabel}
                offAriaLabel={strings.ToggleOffAriaLabel}
                onText={strings.ToggleOnText}
                offText={strings.ToggleOffText}
                onChanged={(checked: boolean) => this.props.onValueChanged(checked.toString())}
            />;
        }
        return element;
    }
    private renderDate = (): JSX.Element => {
        let element: JSX.Element = null;
        const locale: string = Locales[this.props.field.LocaleId];
        moment.locale(locale);
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            let value: string = '';
            if (this.props.field.DisplayFormat === 1) {// date and time
                value = (this.props.value && moment(this.props.value).isValid()) ? moment(this.props.value).format('L HH:mm') : '';
            } else {
                value = (this.props.value && moment(this.props.value).isValid()) ? moment(this.props.value).format('L') : '';
            }
            element = <FieldDateRenderer text={value} />;
        } else {
            if (this.props.field.DisplayFormat === 1) {// date and time
                element = <DateTimePicker
                    {...this.props.value && moment(this.props.value).isValid() ? { value: moment(this.props.value).toDate() } : {}}
                    placeholder={strings.DateFormFieldPlaceholder}
                    formatDate={(date: Date) => (typeof date.toLocaleDateString === 'function') ? date.toLocaleDateString(locale) : ''}
                    timeConvention={TimeConvention.Hours24}
                    onChange={(date) => {
                        if (date) {
                            this.props.onValueChanged(date.toISOString());
                        } else {
                            this.props.onValueChanged('');
                        }
                    }}
                />;
            } else {
                element = <DatePicker
                    {...this.props.value && moment(this.props.value).isValid() ? { value: moment(this.props.value).toDate() } : {}}
                    className={styles.dateFormField}
                    placeholder={strings.DateFormFieldPlaceholder}
                    isRequired={this.props.field.Required}
                    ariaLabel={this.props.field.Title}
                    parseDateFromString={(dateStr?: string) => { return moment(dateStr, 'L').toDate(); }}
                    formatDate={(date: Date) => (typeof date.toLocaleDateString === 'function') ? date.toLocaleDateString(locale) : ''}
                    strings={strings}
                    firstDayOfWeek={this.props.field.FirstDayOfWeek}
                    allowTextInput
                    onSelectDate={(date) => {
                        if (date) {
                            this.props.onValueChanged(date.toISOString());
                        } else {
                            this.props.onValueChanged('');
                        }
                    }}
                />;
            }
        }
        return element;
    }
    private renderChoice = (): JSX.Element => {
        let element: JSX.Element = null;

        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            const value: string = (this.props.value) ? ((typeof this.props.value === 'string') ? this.props.value : this.props.value.join(', ')) : '';
            element = <FieldTextRenderer
                        text={value}
                    />;
        } else {
            if (this.props.field.Type === 'MultiChoice') {
                const options: string[] = this.props.field.MultiChoices;
                let values: any = this.props.value ? this.props.value : [];
                if (!Array.isArray(values)) {
                    values = [values];
                }
                element = <Dropdown
                    title={JSON.stringify(this.props.field) + this.props.value}
                    className={styles.dropDownFormField}
                    options={options.map((option: string) => ({ key: option, text: option }))}
                    selectedKeys={values}
                    placeHolder={strings.MultiChoiceFormFieldPlaceholder}
                    multiSelect
                    onChanged={(item) => this.props.onValueChanged(getUpdatedValue(values, item))}
                />;
            } else {
                // Choice
                const options: any[] = (this.props.field.Required) ? this.props.field.Choices : [''].concat(this.props.field.Choices);
                element = <Dropdown
                    className={styles.dropDownFormField}
                    options={options.map((option: string) => ({ key: option, text: option }))}
                    selectedKey={this.props.value}
                    placeHolder={strings.ChoiceFormFieldPlaceholder}
                    onChanged={(item) => this.props.onValueChanged(item.key.toString())}
                />;
            }
        }
        return element;

        function getUpdatedValue(oldValues: any[], changedItem: IDropdownOption): any[] {
            const changedKey: string = changedItem.key.toString();
            const newValues: any[] = [...oldValues];
            console.log (newValues);
            if (changedItem.selected) {
                // add option if it's checked
                if (newValues.indexOf(changedKey) < 0) { newValues.push(changedKey); }
            } else {
                // remove the option if it's unchecked
                const currIndex: number = newValues.indexOf(changedKey);
                if (currIndex > -1) { newValues.splice(currIndex, 1); }
            }
            return newValues;
        }
    }
    private renderNumber = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            const value: string = (this.props.value) ?
                ((typeof this.props.value === 'string') ?
                    this.props.value
                    :
                    JSON.stringify(this.props.value))
                :
                '';
            element = <FieldTextRenderer
                        text={value}
                    />;
        } else {
            let val: string;
            val = this.props.value ? this.props.value.replace(/[^0-9.,-]+/g, '') : '';
            element = <TextField
                value={val}
                placeholder={strings.NumberFormFieldPlaceholder}
                onChanged={(newValue) => this.props.onValueChanged(newValue)}
                onGetErrorMessage={(value) =>
                    validateNumber (value, this.props.context.pageContext.cultureInfo.currentCultureName)
                }
            />;
        }
        return element;

        function validateNumber(value: string, locale: string): string {
            return isNaN(parseNumber(value, locale))
              ? `${strings.InvalidNumberValue} ${value}`
              : '';
          }

        function parseNumber (value: string, locale: string = navigator.language): number {
            const decimalSperator: string = Intl.NumberFormat(locale).format(1.1).charAt(1);
            // const cleanPattern = new RegExp(`[^-+0-9${ example.charAt( 1 ) }]`, 'g');
            const cleanPattern: RegExp = new RegExp(`[${'\' ,.'.replace(decimalSperator, '')}]`, 'g');
            const cleaned: string = value.replace(cleanPattern, '');
            const normalized: string = cleaned.replace(decimalSperator, '.');
            return Number(normalized);
        }
    }
    private renderAttachments = (): JSX.Element => {
        let element: JSX.Element = null;
        if (this.props.formRenderMethod === 1) { // 1 - display, 2 - edit, 3 - new
            element = <div></div>;
            if (this.props.value) {
                if ((this.props.value.Attachments) && (Array.isArray(this.props.value.Attachments))) {
                    for (let r: number = 0; r < this.props.value.Attachments.length; r++) {
                        element = <div>
                            {this.props.value.Attachments.map ((at: any) => {
                                 return <div className={ardStyles.Attachment}>
                                            <img
                                                alt=''
                                                role='presentation'
                                                src={this._utilities.GetFileImageUrl2(at)}
                                                className={ardStyles.fileIconImg}
                                            />
                                            <Link
                                                target={'_blank'}
                                                style={{marginLeft: '0.5rem'}}
                                                href={at.RedirectUrl ?
                                                        at.RedirectUrl.charAt(0) === '1' ?
                                                            at.RedirectUrl.substring(1)
                                                            : at.RedirectUrl
                                                        :
                                                        this.props.value.UrlPrefix + at.FileName + '?Web=1'}
                                            >
                                                {at.FileName}
                                            </Link>
                                        </div>;
                            })}
                        </div>;
                    }
                }
            }
        } else {
            element =   <div>
                            {this.renderOldAttachments()}
                            <ActionButton
                                data-automation-id='Attach'
                                iconProps={ { iconName: 'Attach' } }
                                onClick = {(ev: React.MouseEvent<any>) => {
                                    ev.preventDefault();
                                    ev.stopPropagation();
                                    const newAttachments: {file: any}[] = this.state.newAttachments.filter(
                                        (attachment) => attachment.file !== null
                                    );
                                    const newAttachmentAddNull: {file: any} = {
                                        file: null
                                    };
                                    newAttachments.push(newAttachmentAddNull);
                                    this.setState({
                                        ...this.state,
                                        newAttachments: newAttachments
                                    }, () => {
                                        this.newAttachmentsRefs.click();
                                    });
                                }}
                            >
                                {strings.AddAttachment}
                            </ActionButton>
                            {
                                (this.state.newAttachments.length > 0) &&
                                <div className={ardStyles.AttachmentsNew}>
                                    {this.state.newAttachments.map((attachment: {file: any}) => {
                                        console.log(attachment);
                                        if (attachment.file === null) {
                                            return this.renderNewAttachmentAdd();
                                        } else {
                                            return this.renderNewAttachments(attachment);
                                        }

                                    })}
                                </div>
                            }
                        </div>;
        }
        return element;
    }
    private renderNewAttachments = (attachment: {file: any}): JSX.Element => {
        return (
            <div className={ardStyles.Attachment}>
                <img
                    alt=''
                    role='presentation'
                    src={this._utilities.GetFileImageUrl1(attachment.file)}
                    className={ardStyles.fileIconImg}
                />
                <Link
                    style={{marginLeft: '0.5rem'}}
                >
                    {attachment.file.name}
                </Link>
                <IconButton
                    iconProps={ { iconName: 'Cancel' }}
                    title={strings.Cancel}
                    ariaLabel={strings.Cancel}
                    onClick = {() => {
                        const newAttachments: {file: any}[] = this.state.newAttachments.filter(
                            (attach) => attach.file.name !== attachment.file.name
                        );
                        this.setState({
                            ...this.state,
                            newAttachments: newAttachments
                        });
                    }}
                />
            </div>
        );
    }
    private renderOldAttachments = (): JSX.Element => {
        return (
            <div className={ardStyles.AttachmentsNew}>
                {this.props.value ?
                this.props.value.Attachments ?
                    (this.props.value.Attachments !== '') ?
                    this.props.value.Attachments
                        .filter((del: any) => {
                            return !(this.state.deleteAttachments.indexOf(del.FileName) > -1);
                        })
                        .map ((at: any) => {
                            return <div className={ardStyles.Attachment}>
                                        <img
                                            alt=''
                                            role='presentation'
                                            src={this._utilities.GetFileImageUrl2(at)}
                                            className={ardStyles.fileIconImg}
                                        />
                                        <Link
                                            target={'_blank'}
                                            style={{marginLeft: '0.5rem'}}
                                            href={at.RedirectUrl ?
                                                    at.RedirectUrl.charAt(0) === '1' ?
                                                        at.RedirectUrl.substring(1)
                                                        : at.RedirectUrl
                                                    :
                                                    this.props.value.UrlPrefix + at.FileName + '?Web=1'}
                                        >
                                            {at.FileName}
                                        </Link>
                                        <IconButton
                                            iconProps={ { iconName: 'Delete' }}
                                            title={strings.Delete}
                                            ariaLabel={strings.Delete}
                                            onClick = {() => {
                                                const deleteA: any[] = [...this.state.deleteAttachments];
                                                deleteA.push(at.FileName);
                                                this.setState({
                                                    ...this.state,
                                                    deleteAttachments: [...deleteA]
                                                });
                                            }}
                                        />
                                    </div>;
                        })
                        :
                        null
                    :
                    null
                    :
                    null
                }

            </div>);
    }
    private renderNewAttachmentAdd = (): JSX.Element => {
        if (this.newAttachmentsRefs === null) {
            return (
                <input
                    style={{ display: 'none' }}
                    type='file'
                    onChange={(e) => this.addAttachment(e)}
                    ref={(newRef) => {
                            this.newAttachmentsRefs = newRef;
                        }
                    }
                />
            );
        } else {
            this.newAttachmentsRefs = null;
            // we had cancel button on input file
            const newAttachments: {file: any}[] = this.state.newAttachments.filter(
                (attachment) => attachment.file !== null
            );
            const newAttachment: {file: any} = {
                file: null
            };
            newAttachments.push(newAttachment);
            this.setState({
                ...this.state,
                newAttachments: newAttachments
            });
            return null;
        }
    }
    private addAttachment = (e: React.ChangeEvent<HTMLInputElement>) => {
        const reader: FileReader = new FileReader();
        const file: File = e.target.files[0];
        reader.onloadend = () => {
            const newAttachments: {file: any}[] = [...this.state.newAttachments];
            const obj: {
                file: any;
            } = newAttachments.find(f => f.file == null);
            obj.file = file;
            this.setState({
                ...this.state,
                newAttachments: newAttachments
            });
        };
        reader.readAsDataURL(file);
    }
    private renderWaiting = () => {
        return this.props.isSaving ?
          <div>
            <Overlay isDarkThemed={false} className={ardStyles.overlay}>
              <Spinner size={SpinnerSize.medium} />
            </Overlay>
          </div>
          :
          null;
      }
}