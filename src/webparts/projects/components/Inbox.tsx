import * as React from 'react';
import { IInboxProps } from './IInboxProps';
import { PnPService } from '../services/PnPService';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, TextField } from 'office-ui-fabric-react';
import { IDetailsListItem, IInboxState} from './IInboxState';

export default class Inbox extends React.Component<IInboxProps, IInboxState> {

    private _inboxresults: Array<IDetailsListItem>;
    private _columns: IColumn[];
    private _pnpService;
    constructor(props) {
        super(props);
        this._pnpService = new PnPService(this.props.context);
        this._columns = [
            { key: 'column1', name: 'Codigo', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column2', name: 'Titulo', fieldName: 'value', minWidth: 100, maxWidth: 200, isResizable: true },
          ];
          this.state = {
            url: this.props.context.pageContext.web.absoluteUrl,
            items: this._inboxresults
        };
        this._inboxresults = [];
        this._pnpService.getTipoDeDocumentos(this.state.url).then(typeDocs => {
            typeDocs.forEach(typeDoc => {
              const newResult: IDetailsListItem = {
                key: typeDoc["Codigo"],
                name: typeDoc["Codigo"],
                value: typeDoc["Title"]
              };
              this._inboxresults.push(newResult);
            });
        });
    }

    public componentDidMount() {
        this.setState({
            url: this.props.context.pageContext.web.absoluteUrl,
            items: this._inboxresults
        });
    }

    private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, name: string): void => {
        this.setState({
            items: name ? this._inboxresults.filter(i => i.name.toLowerCase().indexOf(name) > -1) : this._inboxresults
        });
    };
    public render(): React.ReactElement<IInboxProps> {
        return (
            <div>
                <div>
                    <TextField label="Id. Bandeja de Entrada" placeholder= "Buscar por Codigo" onChange={this._onFilter.bind(this)} />
                </div>
                <DetailsList
                    items={this.state.items}
                    columns={this._columns}
                    selectionMode={SelectionMode.none}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                />
            </div>
        );
    }
}
