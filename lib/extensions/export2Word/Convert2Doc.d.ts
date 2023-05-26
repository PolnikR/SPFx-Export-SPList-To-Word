import { IListItemsResponse } from 'spfxhelper/dist/SPFxHelper/Props/ISPListProps';
import { SPHttpClient } from '@microsoft/sp-http';
declare class Convert2Doc {
    private _logSource;
    private _webURL;
    private _client;
    private listName;
    private response;
    private currentViewFields;
    private listFieldDetails;
    private cisloZiadanky;
    private _listOperations;
    private readonly listOperations;
    private _commonOperations;
    private readonly commonOperations;
    private _fieldOperations;
    private readonly fieldOperations;
    /**
    * Creates the select query for retreiving the records
    */
    private readonly createSelectQuery;
    constructor(spHttp: SPHttpClient, webUrl: string, logSource: string, listName: string);
    /**
     * Creates the document for the export for the current list with current view fields
     */
    createDocument(): Promise<void>;
    /**
     * Returns all the items in a list
     * @param listName list name on which query needs to be performed
     */
    private getItems;
    /**
     * Recursively gets all the items in the list
     * @param nextLink
     */
    private getAllItems;
    /**
     * Returns the fields in the current View
     * @param listName
     */
    getCurrentViewFields(): Promise<void>;
    /**
    * Validates the column types for QnA format
    * @param listName listname
    */
    private validateColumnTypes;
    /**
     * Generates the table format for the output
     * @param items
     */
    generateTableFormat(items: IListItemsResponse): void;
    /**
     * Generates the QnA format for the output
     * @param items
     */
    generateQnAFormat(items: IListItemsResponse): void;
    /**
    * Generates the document and download it
    * @param sourceHTML
    */
    generateDocument(sourceHTML: string, cisloZiadanky: string): void;
}
export { Convert2Doc };
//# sourceMappingURL=Convert2Doc.d.ts.map