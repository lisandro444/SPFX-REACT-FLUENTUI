import { INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { Web } from '@pnp/sp/webs';
import { ICamlQuery } from '@pnp/sp/lists';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// import { IWebPartContext } from '@microsoft/sp-webpart-base';  
import { ISPSearchResult } from '../components/ISPSearchResult';
import { ISearchResults, ICells, ICellValue } from './ISearchService';

export class PnPService {
    private _context;
    private _siteUrl;
    private docsLibraries: any = ["Analisis", "Adquisiciones", "Contrato", "Financieros"];
    constructor(context: BaseWebPartContext) {
        this._context = context;
        this._siteUrl = context.pageContext.web.absoluteUrl;
    }

    public async getProjectsWithSite(url): Promise<any[]> {
        try {
            let initialweb = Web(url);
            const caml: ICamlQuery = {
                ViewXml: "<View><Query><Where><IsNotNull><FieldRef Name='Link'/></IsNotNull></Where></Query></View>",
            };
            let items = await initialweb.lists.getByTitle("Proyectos").getItemsByCAMLQuery(caml);
            console.log("CAML Query");
            console.log(items);
            return items;
        }
        catch (error) {
            console.log("getProjects: " + error);
            return null;
        }
    }

    public async getTipoDeDocumentos(url): Promise<any[]> {
        try {
            let initialweb = Web(url);
            let items = await initialweb.lists.getByTitle("Codificación").items.get();
            console.log("Tipo de decumentos: ");
            console.log(items);
            return items;
        }
        catch (error) {
            console.log("getTipoDeDocumentos: " + error);
            return null;
        }
    }

    public async getInbox(url): Promise<any[]> {
        try {
            let initialweb = Web(url);
            let items = await initialweb.lists.getByTitle("Bandeja de Entrada").items.get();
            console.log("Inbox: ");
            console.log(items);
            return items;
        }
        catch (error) {
            console.log("getInbox: " + error);
            return null;
        }
    }

    public async getInboxById(Id): Promise<any> {
        try {
            let initialweb = Web(this._siteUrl);
            let docs = await initialweb.lists.getByTitle("Bandeja de Entrada").items.getById(Id);
            console.log("Doc asociado: ");
            console.log(docs);
            return docs;
        }
        catch (error) {
            console.log("getInboxById: " + error);
            return null;
        }
    }

    // Get Documents
    public GetAllDocumentsFromProject(url): any[] {
        let resultsDocs: any[] = [];
        try {
            let initialweb = Web(url);
            initialweb.lists.get().then(async data => {
                for (var i = 0; i < data.length; i++) {
                    if (data[i].BaseTemplate == 101 && this.docsLibraries.indexOf(data[i].Title) > -1) {
                        let docs = await initialweb.lists.getByTitle(data[i].Title).items.get();
                        let docsFiltered = docs.filter(x => x.TipoDocumento != null);
                        docsFiltered.forEach(doc => {
                            resultsDocs.push(doc);
                        });
                    }
                }
                console.log("Documentos desde el servicio.........................");
                console.log(resultsDocs);
                return resultsDocs;
            }).catch(data => { console.log(data); });
        } catch (error) {
            console.log("GetAllDocumentsFromProject: " + error);
            return error;
        }
    }

    private getTermNameById(id: string) {
        //const stores = await taxonomy.termStores.get();
        // const store: ITermStore = taxonomy.termStores.getByName("Taxonomy_v5o/SbcTE2cegwO2dtAN9l==");

        // const group: ITermGroup = store.getTermGroupById("0ba6845c-1468-4ec5-a5a8-718f1fb05431");
    }

    // Search Logic

    public async getSearchResults(query: string): Promise<ISPSearchResult[]> {

        let url: string = this._context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=" + query;
        console.log(url);
        return new Promise<ISPSearchResult[]>((resolve, reject) => {
            // Do an Ajax call to receive the search results  
            this._getSearchData(url).then((res: ISearchResults) => {
                let searchResp: ISPSearchResult[] = [];

                // Check if there was an error  
                if (typeof res["odata.error"] !== "undefined") {
                    if (typeof res["odata.error"]["message"] !== "undefined") {
                        Promise.reject(res["odata.error"]["message"].value);
                        return;
                    }
                }

                if (!this._isNull(res)) {
                    const fields: string = "Title,Path,Description";

                    // Retrieve all the table rows  
                    if (typeof res.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {
                        if (typeof res.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {
                            searchResp = this._setSearchResults(res.PrimaryQueryResult.RelevantResults.Table.Rows, fields);
                        }
                    }
                }

                // Return the retrieved result set  
                resolve(searchResp);
            });
        });
    }

    /** 
    * Retrieve the results from the search API 
    * 
    * @param url 
    */
    private _getSearchData(url: string): Promise<ISearchResults> {
        return this._context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
            headers: {
                'odata-version': '3.0'
            }
        }).then((res: SPHttpClientResponse) => {
            return res.json();
        }).catch(error => {
            return Promise.reject(JSON.stringify(error));
        });
    }

    /** 
     * Set the current set of search results 
     * 
     * @param crntResults 
     * @param fields 
     */
    private _setSearchResults(crntResults: ICells[], fields: string): any[] {
        const temp: any[] = [];

        if (crntResults.length > 0) {
            const flds: string[] = fields.toLowerCase().split(',');

            crntResults.forEach((result) => {
                // Create a temp value  
                var val: Object = {}

                result.Cells.forEach((cell: ICellValue) => {
                    if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {
                        // Add key and value to temp value  
                        val[cell.Key] = cell.Value;
                    }
                });

                // Push this to the temp array  
                temp.push(val);
            });
        }

        return temp;
    }

    /** 
     * Check if the value is null or undefined 
     * 
     * @param value 
     */
    private _isNull(value: any): boolean {
        return value === null || typeof value === "undefined";
    }
} 