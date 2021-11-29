declare interface IMFormWebPartStrings {
  Yes: string;
  No: string;
  Save: string;
  Cancel: string;
  Delete: string;
  ToggleOnAriaLabel: string;
  ToggleOffAriaLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
  DateFormFieldPlaceholder: string;
  UrlFormFieldPlaceholder: string;
  UrlDescFormFieldPlaceholder: string;
  TextFormFieldPlaceholder: string;
  UserFormFieldPlaceholder: string;
  UsersFormFieldPlaceholder: string;
  ChoiceFormFieldPlaceholder: string;
  MultiChoiceFormFieldPlaceholder: string;
  NumberFormFieldPlaceholder: string;
  AddAttachment: string;
  ItemSavedSuccessfully: string;
  InvalidNumberValue: string;
  months: string[];
  shortMonths: string[];
  days: string[];
  shortDays: string[];
  goToToday: string;
  prevMonthAriaLabel?: string;
  nextMonthAriaLabel?: string;
  prevYearAriaLabel?: string;
  nextYearAriaLabel?: string;
  Errors: {
    ErrorWebAccessDenied: string;
    ErrorOnLoadingApp: string;
    ErrorWebNotFound: string;
    RequiredValueMessage: string;
    ErrorOnSavingListItem: string;
    ErrorDuplicateAttachment: string;
    FieldsErrorOnSaving: string;
    ErrorNoFields: string;
  };
}

declare module 'MFormWebPartStrings' {
  const strings: IMFormWebPartStrings;
  export = strings;
}
