import * as React from 'react';
import styles from './SiteProvisioningTool.module.scss';
import { ISiteProvisioningToolProps } from './ISiteProvisioningToolProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";

import MainForm, { IFormData } from './custom-components/MainForm';
import DocumentCardCreateSite from './DocumentCardCreateSite';
export interface ISiteProvisioningToolState {
  loadForm: boolean;
  showCurrentStatus: boolean;
  currentStatus: string;
  currentCreatedSiteUrl : string;
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

const graphUrl = 'https://graph.microsoft.com/v1.0/groups';

export default class SiteProvisioningTool extends React.Component<ISiteProvisioningToolProps, ISiteProvisioningToolState> {
  constructor(props: ISiteProvisioningToolProps, state: ISiteProvisioningToolState) {
    super(props);
    this.state = {
      loadForm: false,
      showCurrentStatus: false,
      currentStatus: '',
      currentCreatedSiteUrl : ''
    };
    this.loadForm = this.loadForm.bind(this);
    this.createSiteCollection = this.createSiteCollection.bind(this);

  }
  // called with the Create Site button is cliecked
  loadForm(event: any) {
    this.setState({ loadForm: true });
  }
  createSiteCollection(formData: IFormData) {
    this.setState({showCurrentStatus : true,currentStatus : "Provisioning your site"});
    var self = this;
    const siteCreationBody: string = JSON.stringify(
      {
        description: formData.siteDescription,
        displayName: formData.siteName,
        groupTypes: [
          "Unified"
        ],
        mailEnabled: false,
        mailNickname: formData.groupEmailAddress,
        securityEnabled: formData.privacyOptions.key == 'Private' ? true : false
      });
    //create the site collection using graph api
    self.createNewSiteCollectionUsingGraph(self, siteCreationBody)
      .then(response => {
        return response.json();
      }).then(data => {
        console.log(data);
        //get root site collection id from group id
        self.setState({showCurrentStatus : true,currentStatus : "Created Site Collection, looking for the site id..."});
        var groupId = data.id;
        setTimeout(function () {
          self.getSiteCollectionIdFromGroupId(self, groupId)
            .then(response => {
              return response.json();
            })
            .then(data => {
              self.setState({showCurrentStatus : true,currentStatus : "Got the Site Collection ID"});
              var siteCollectionId = data.id;
              self.setState({currentCreatedSiteUrl : data.webUrl});
              var listCreationBody = JSON.stringify({
                displayName: "Books",
                columns: [
                  {
                    name: "Author",
                    text: {}
                  },
                  {
                    name: "PageCount",
                    number: {}
                  }
                ],
                list: {
                  "template": "documentLibrary"
                }
              });
              self.createDocumentLibraries(self, siteCollectionId, listCreationBody)
                .then(response => {
                  return response.json();
                })
                .then(data => {
                  self.setState({showCurrentStatus : true,currentStatus : "Your brand new team site has been created"});
                  Object.assign(document.createElement('a'), { target: '_blank', href: self.state.currentCreatedSiteUrl}).click();
                });
            });
        }, 10000);


      });
  }

  private createDocumentLibraries(self: this, siteCollectionId: any, listCreationBody: string) {
    return self.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .post(`https://graph.microsoft.com/v1.0/sites/${siteCollectionId}/lists`, AadHttpClient.configurations.v1, {
            headers: {
              'Content-type': 'application/json',
            },
            body: listCreationBody
          });
      });
  }

  private getSiteCollectionIdFromGroupId(self: this, groupId: any) {
    return self.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .get(`https://graph.microsoft.com/v1.0/groups/${groupId}/sites/root`, AadHttpClient.configurations.v1
          );
      });
  }

  private createNewSiteCollectionUsingGraph(self: this, siteCreationBody: string) {
    return self.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .post(`https://graph.microsoft.com/v1.0/groups`, AadHttpClient.configurations.v1, {
            headers: {
              'Content-type': 'application/json',
            },
            body: siteCreationBody
          });
      });
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
    return this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
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
              <span className={styles.title}>Create a site : {this.props.context.pageContext.web.title}</span>
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

        </div> : <MainForm createSiteCollection={this.createSiteCollection} spContext={this.props.context} currentStatus={this.state.currentStatus} showCurrentStatus={this.state.showCurrentStatus} />}

      </div>
    );
  }
}
