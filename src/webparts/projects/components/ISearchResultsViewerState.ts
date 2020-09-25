import {ISPSearchResult} from './ISPSearchResult';  
  
export interface ISearchResultsViewerState {  
    status: string;  
    searchText: string;  
    items: ISPSearchResult[];  
}  