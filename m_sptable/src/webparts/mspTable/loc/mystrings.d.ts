declare interface IMspTableWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PropertyPane: {
    ListFieldLabel: string;
    EditRefinersLabel: string;
    RefinersFieldDescription: string;
    RefinersFieldLabel: string;
    ShowExpanded: string;
    showValueFilter: string;
    Search: string;
    debugTitle: string;
    debugLabel: string;
    PropertyEdit: string;
    Templates: {
      RefinementItemTemplateLabel: string;
      MutliValueRefinementItemTemplateLabel: string;
      PersonaRefinementItemLabel: string;
      FixedDateRangeRefinementItemLabel: string;
      RefinerSortTypeLabel: string;
      RefinerSortTypeByNumberOfResults: string;
      RefinerSortTypeAlphabetical: string;
      RefinerSortTypeSortOrderLabel: string;
      RefinerSortTypeSortDirectionAscending: string;
      RefinerSortTypeSortDirectionDescending: string;
      FilterInternalName: string;
      FilterMode: string;
    }
  }
}

declare module 'MspTableWebPartStrings' {
  const strings: IMspTableWebPartStrings;
  export = strings;
}
