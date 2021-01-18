export interface IFilterOptions {
    title: string;
    fieldName: string;
    options: {
        value: string;
        text: string;
        selected: boolean;
    }[];
}
