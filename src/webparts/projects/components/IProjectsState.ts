import { IDropdownOption } from "office-ui-fabric-react";

export interface IProjectsPropsState {
    url:string;
    projects: Array<IDropdownOption>;
    typeDocs: Array<IDropdownOption>;
}