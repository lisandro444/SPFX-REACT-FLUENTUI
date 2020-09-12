import * as React from 'react';
import styles from './Projects.module.scss';
import { IProjectsProps } from './IProjectsProps';
import { PnPService } from '../services/PnPService';
import { IProjectsPropsState } from './IProjectsState';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IProject } from '../Entities/Project';

export default class Projects extends React.Component<IProjectsProps, IProjectsPropsState> {

  private _projects: Array<IDropdownOption>;
  private _pnpService;
  constructor(props) {
    super(props);
    this._pnpService = new PnPService(this.props.context);
    this.state = {
      url: this.props.context.pageContext.web.absoluteUrl,
      projects: this._projects
    };
  }

  public componentDidMount() {
    this._projects = new Array<IDropdownOption>();
    this._pnpService.getProjectsWithSite(this.state.url).then(projects => {
        projects.forEach(project => {

          const newProject: IDropdownOption = {
            key: project["CodigoProyecto"],
            text: project["Title"]
          };
          this._projects.push(newProject);
        });
    });

    this.setState({
      url: this.props.context.pageContext.web.absoluteUrl,
      projects:this._projects
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
            <span className={styles.title}>Buscador Tipo de Documentos</span>
              <Dropdown
                placeholder="Select project"
                label="Projects"
                multiSelect
                options={this.state.projects}
                styles={dropdownStyles}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
