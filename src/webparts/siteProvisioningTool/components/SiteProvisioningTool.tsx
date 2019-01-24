import * as React from 'react';
import styles from './SiteProvisioningTool.module.scss';
import { ISiteProvisioningToolProps } from './ISiteProvisioningToolProps';
import { escape } from '@microsoft/sp-lodash-subset';

import MainForm from './custom-components/MainForm';
import DocumentCardCreateSite from './DocumentCardCreateSite';
export interface ISiteProvisioningToolState {
  loadForm: boolean;
}

import MockHttpClient from './MockHttpClient';

export interface ISPLists {
  value: ISPList[];
 }
 
 export interface ISPList {
  Title: string;
  Id: string;
 }

import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

export default class SiteProvisioningTool extends React.Component<ISiteProvisioningToolProps, ISiteProvisioningToolState> {
  constructor(props: ISiteProvisioningToolProps, state: ISiteProvisioningToolState) {
    super(props);
    this.state = {
      loadForm: false
    };
    this.loadForm = this.loadForm.bind(this);
    this.createSiteCollection = this.createSiteCollection.bind(this);

  }
  // called with the Create Site button is cliecked
  loadForm(event: any) {
    this.setState({ loadForm: true });
  }
  createSiteCollection() {

  }

  componentDidMount() {
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        console.log(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          console.log(response.value);
        });
    }
  }
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }
  public render(): React.ReactElement<ISiteProvisioningToolProps> {
    return (

      <div className={styles.siteProvisioningTool}>
        {(this.state.loadForm == false) ? <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Create a site : </span>
            </div>

          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              {/* <button onClick={this.loadForm} className={styles.button}>
                <span className={styles.label}>Create Site</span>
              </button> */}
              <DocumentCardCreateSite formLoadClick={this.loadForm} />
            </div>

          </div>

        </div> : <MainForm createSiteCollection={this.createSiteCollection} spContext={this.context} />}

      </div>
    );
  }
}
