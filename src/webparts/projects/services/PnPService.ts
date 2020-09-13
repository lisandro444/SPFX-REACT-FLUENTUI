import { INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { Web } from '@pnp/sp/webs';
import { ICamlQuery } from '@pnp/sp/lists';
import { IProject } from '../Entities/Project';

export class PnPService {
    private _context;
    private _siteUrl;
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

}