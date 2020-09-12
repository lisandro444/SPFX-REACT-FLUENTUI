import * as React from 'react';
import styles from './Projects.module.scss';
import { IProjectsProps } from './IProjectsProps';
import { PnPService } from '../services/PnPService';
import { IProjectsPropsState } from './IProjectsState';

export default class Projects extends React.Component<IProjectsProps, IProjectsPropsState> {

  private _pnpService;
  constructor(props) {
    super(props);
    this._pnpService = new PnPService(this.props.context);
    this.state = {
      url: this.props.context.pageContext.web.absoluteUrl,
    };
  }

  public componentDidMount(){
    this._pnpService.getProjectsWithSite(this.state.url).then(items=>{
    });
  }

  public render(): React.ReactElement<IProjectsProps> {
    return (
      <div className={ styles.projects }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>   
            </div>
          </div>
        </div>
      </div>
    );
  }
}
