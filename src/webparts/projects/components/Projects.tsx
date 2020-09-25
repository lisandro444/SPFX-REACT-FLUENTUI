import * as React from 'react';
import styles from './Projects.module.scss';
import { IProjectsProps } from './IProjectsProps';
import { PnPService } from '../services/PnPService';
import { IProjectsPropsState } from './IProjectsState';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IProject } from '../Entities/Project';
import { TextField } from 'office-ui-fabric-react';
import Inbox from './Inbox';

export default class Projects extends React.Component<IProjectsProps, IProjectsPropsState> {

  private _projects: Array<IDropdownOption>;
  private _typeDocs: Array<IDropdownOption>;
  private _pnpService;
  constructor(props) {
    super(props);
    this._pnpService = new PnPService(this.props.context);
    this.state = {
      url: this.props.context.pageContext.web.absoluteUrl,
      projects: this._projects,
      typeDocs: this._typeDocs,
      status: 'Ready',  
      searchText: '',  
      items: [],
    };
  }

  public componentDidMount() {
    this._projects = new Array<IDropdownOption>();
    this._typeDocs = new Array<IDropdownOption>();
    this._pnpService.getProjectsWithSite(this.state.url).then(projects => {
        projects.forEach(project => {

          const newProject: IDropdownOption = {
            key: project["CodigoProyecto"],
            text: project["Title"]
          };
          this._projects.push(newProject);
        });
    });
    this._pnpService.getTipoDeDocumentos(this.state.url).then(typeDocs => {
      typeDocs.forEach(typeDoc => {
        const newTypeDoc: IDropdownOption = {
          key: typeDoc["Codigo"],
          text: typeDoc["Title"]
        };
        this._typeDocs.push(newTypeDoc);
      });
  });

  const queryText = "sharepoint";
  this._pnpService.getSearchResults(queryText).then(result => {
    console.log(result.items);
  });

    this.setState({
      url: this.props.context.pageContext.web.absoluteUrl,
      projects:this._projects,
      typeDocs:this._typeDocs
    });
  }

  public render(): React.ReactElement<IProjectsProps> {
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 300 },
    };
    return (
      <div className={styles.projects}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
            <span className={styles.title}>Reporte</span>
              <Dropdown
                placeholder="Select project"
                label="Projects"
                multiSelect
                options={this.state.projects}
                styles={dropdownStyles}
              />
              <Dropdown
                placeholder="Select Tipo de Documento"
                label="Tipo de Documentos"
                multiSelect
                options={this.state.typeDocs}
                styles={dropdownStyles}
              />
              <Inbox context={this.props.context} /> 
            </div>
          </div>
        </div>
      </div>
    );
  }
}
