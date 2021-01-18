define([], function() {
    return {
      Templates: {
        DateFromLabel: 'Od',
        DateTolabel: 'Do',
        DatePickerStrings: {
          months: ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September', 'Október', 'November', 'December'],
          shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'Máj', 'Jún', 'Júl', 'Aug', 'Sep', 'Okt', 'Nov', 'Dec'],
          days: ['Nedeľa', 'Pondelok', 'Utorok', 'Streda', 'Štvrtok', 'Piatok', 'Sobota'],
          shortDays: ['Ne', 'Po', 'Ut', 'St', 'Št', 'Pi', 'So'],
          goToToday: 'Dnes',
          prevMonthAriaLabel: 'Predchádzajúci mesiac',
          nextMonthAriaLabel: 'Nasledujúci mesiac',
          prevYearAriaLabel: 'Prechádzajúci rok',
          nextYearAriaLabel: 'Nasledujúci rok',
          closeButtonAriaLabel: 'Zatvoriť',
          isRequiredErrorMessage: 'Tento dátum je povinný.',
          invalidInputErrorMessage: 'Nesprávny formát dátumu.'
        },
      },
      Filters: {
        ClearFiltersLabel: 'Zrušiť',
        ApplyFiltersLabel: 'Aplikovať',
        ShowAll: 'Zobraziť všetko',
        FilterPlacehoder: 'Filter',
        RemoveAllFiltersLabel: 'Vymazať všetky filtre',
        FilterPanelTitle: 'Filtre',
        FilterPanelClose: 'Zatvoriť',
      },
      Table: {
        ResultsCount: 'Počet',
        SearchPlaceholder: 'Vyhľadávanie',
        ChangeView: 'Zmeniť zobrazenie',
        filterText: 'Filtre',
      },
      Errors: {
        ErrorLoadingData: 'Chyba pri získavaní údajov: ',
      },
    }
  });