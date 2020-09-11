import { INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@pnp/sp';

export class PnPService {
    private _context;
    constructor(context: BaseWebPartContext) {
        this._context = context;
    }

    public async getWeb(): Promise<any> {
        try {
            const w = await sp.web.select("DICELTeam")();
            return w.Title;
        }
        catch (error) {
            console.log("Get WEb: " + error);
            return null;
        }
    }

}