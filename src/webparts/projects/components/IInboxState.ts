import { IColumn, IDropdownOption } from "office-ui-fabric-react";

export interface IInboxState {
    url:string;
    items:IDetailsListItem[];
}

export interface IDetailsListItem {
    key: string;
    name: string; 
    value:string;
  }