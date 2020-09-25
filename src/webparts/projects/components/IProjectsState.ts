import { IDropdownOption } from "office-ui-fabric-react";
import { ISPSearchResult } from "./ISPSearchResult";

export interface IProjectsPropsState {
    url:string;
    projects: Array<IDropdownOption>;
    typeDocs: Array<IDropdownOption>;
    status: string;  
    searchText: string;  
    items: ISPSearchResult[];
}