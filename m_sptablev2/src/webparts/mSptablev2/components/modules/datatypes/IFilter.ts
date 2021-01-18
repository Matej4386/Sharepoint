export interface IFilter {
    title: string;
    fieldName: string;
    options: {
        value: string;
        text: string;
        selected: boolean;
        executed: boolean;
    }[];
}
