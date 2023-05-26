import { BaseListViewCommandSet, IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
/**
 *
 */
export interface IExport2WordCommandSetProperties {
    listItems: [{
        "ID": "";
        "Kam": "";
    }];
    ID: string;
}
export default class Export2WordCommandSet extends BaseListViewCommandSet<IExport2WordCommandSetProperties> {
    onInit(): Promise<void>;
    onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void;
    onExecute(event: IListViewCommandSetExecuteEventParameters): void;
    /**
     * Creates the documents for the selected items only
     * @param event
     * @param cnvrt2docx
     */
    returnID(): string;
    private getUserProperties;
    private dateConvert;
    private createDocumentSelectedItems;
}
//# sourceMappingURL=Export2WordCommandSet.d.ts.map