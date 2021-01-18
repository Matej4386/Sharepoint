declare interface IMspTableStrings {
  Templates: {
    DateFromLabel: string;
    DateTolabel: string;
    DatePickerStrings: {
      months: string[];
      shortMonths: string[];
      days: string[];
      shortDays: string[];
      goToToday: string;
      prevMonthAriaLabel: string;
      nextMonthAriaLabel: string;
      prevYearAriaLabel: string;
      nextYearAriaLabel: string;
      closeButtonAriaLabel: string;
      isRequiredErrorMessage: string;
      invalidInputErrorMessage: string;
    };
  };
  Filters: {
    ClearFiltersLabel: string;
    ApplyFiltersLabel: string;
    ShowAll: string;
    FilterPlacehoder: string;
    RemoveAllFiltersLabel: string;
    FilterPanelTitle: string;
    FilterPanelClose: string;
  };
  Table: {
    ResultsCount: string;
    SearchPlaceholder: string;
    ChangeView: string;
    filterText: string;
  };
  Errors: {
    ErrorLoadingData: string;
  };
}

declare module 'MspTableStrings' {
  const strings: IMspTableStrings;
  export = strings;
}
