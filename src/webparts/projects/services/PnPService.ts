import { INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@pnp/sp';

export class PnPService {
    private _context;
    constructor(context: BaseWebPartContext) {
        this._context = context;
    }

    public async getWeb(): Promise<any> {
        try {
            const w = await sp.web.select("Gerencia de Operaciones")();
            return w.Title;
        }
        catch (error) {
            console.log("Get WEb: " + error);
            return null;
        }
    }

    public async getProjectsWithSite(): Promise<any> {
        try {
            // get all the items from a list
            const items: any[] = await sp.web.lists.getByTitle("Proyectos").items.get();
            console.log(items);
        }
        catch (error) {
            console.log("getProjects: " + error);
            return null;
        }
    }

}