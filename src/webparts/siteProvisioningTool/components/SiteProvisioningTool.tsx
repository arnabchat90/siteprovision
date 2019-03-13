import * as React from 'react';
import styles from './SiteProvisioningTool.module.scss';
import { ISiteProvisioningToolProps } from './ISiteProvisioningToolProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');

import MainForm, { IFormData } from './custom-components/MainForm';
import DocumentCardCreateSite from './DocumentCardCreateSite';
export interface ISiteProvisioningToolState {
  loadForm: boolean;
  showCurrentStatus: boolean;
  currentStatus: string;
  currentCreatedSiteUrl: string;
  loadingLists: boolean;
  listTitles: any;
  error: any;
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
import { setLanguage } from '@uifabric/utilities/lib';

const graphUrl = 'https://graph.microsoft.com/v1.0/groups';

function onQueryFailed(sender, args) {

  alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

function createFolder(list, folderUrl) {
  var createFolderInternal = function (parentFolder, folderUrl) {
    var ctx = parentFolder.get_context();
    var folderNames = folderUrl.split("/");
    var folderName = folderNames.shift();
    var folder = parentFolder.get_folders().add(folderName);
    ctx.load(folder);
    return executeQuery(ctx)
      .then(function () {
        if (folderNames.length > 0) {
          return createFolderInternal(folder, folderNames.join("/"));
        }
        return folder;
      });
  };
  return createFolderInternal(list.get_rootFolder(), folderUrl);
}

function executeQuery(context) {
  return new Promise(function (resolve, reject) {
    context.executeQueryAsync(function () {
      resolve();
    }, function (sender, args) {
      reject(args);
    });
  });
}
export default class SiteProvisioningTool extends React.Component<ISiteProvisioningToolProps, ISiteProvisioningToolState> {
  constructor(props: ISiteProvisioningToolProps, state: ISiteProvisioningToolState) {
    super(props);
    this.state = {
      loadForm: false,
      showCurrentStatus: false,
      currentStatus: '',
      currentCreatedSiteUrl: '',
      loadingLists: false,
      listTitles: [],
      error: ''
    };
    this.loadForm = this.loadForm.bind(this);
    this.createSiteCollection = this.createSiteCollection.bind(this);
    this.getListsTitles = this.getListsTitles.bind(this);
    this.createDocumentLibrariesJSOM = this.createDocumentLibrariesJSOM.bind(this);
    this.readLibraryConfigurationList = this.readLibraryConfigurationList.bind(this);
    this.createFolderStructure = this.createFolderStructure.bind(this);
  }

  private getListsTitles(siteUrl): void {
    this.setState({
      loadingLists: true,
      listTitles: [],
      error: null
    });

    const context: SP.ClientContext = new SP.ClientContext(siteUrl);
    const lists: SP.ListCollection = context.get_web().get_lists();
    context.load(lists, 'Include(Title)');
    context.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
      const listEnumerator: IEnumerator<SP.List> = lists.getEnumerator();

      const titles: string[] = ['Current Site Collection - ' + siteUrl];
      while (listEnumerator.moveNext()) {
        const list: SP.List = listEnumerator.get_current();
        titles.push(list.get_title());
      }
      console.log(titles);
      this.setState((prevState: ISiteProvisioningToolState, props: ISiteProvisioningToolProps): ISiteProvisioningToolState => {
        prevState.listTitles = titles;
        prevState.loadingLists = false;
        return prevState;
      });
    }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
      this.setState({
        loadingLists: false,
        listTitles: [],
        error: args.get_message()
      });
    });
  }


  private readLibraryConfigurationList(clientContext: SP.ClientContext, libName: string): Promise<any> {
    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      var self = this;
      var oList = clientContext.get_web().get_lists().getByTitle(libName);
      var query = SP.CamlQuery.createAllItemsQuery();
      var collListItem = oList.getItems(query);
      clientContext.load(collListItem);
      clientContext.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
        const listEnumerator: IEnumerator<SP.ListItem> = collListItem.getEnumerator();
        let itemObjects = [];

        while (listEnumerator.moveNext()) {
          const listItem: SP.ListItem = listEnumerator.get_current();
          var folderArr = [];
          // for (var i = 1; i <= 10; i++) {
          //   folderArr[i] = "";
          // }
          var folderTree = "";
          for (var i = 1; i <= 10; i++) {
            if (listItem.get_item('Level' + i) !== "" && listItem.get_item('Level' + i) !== null && listItem.get_item('Level' + i) !== undefined) {
              folderArr.push(listItem.get_item('Level' + i));
            }
          }

          if (folderArr.length > 0) {
            folderTree = folderArr.join('/');
          }
          itemObjects.push({
            LibraryName: listItem.get_item('Title'),
            FolderStructure: folderTree
          });

        }
        resolve(itemObjects);

      }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
        reject(args.get_message());
        self.setState({
          loadingLists: false,
          listTitles: [],
          error: args.get_message()
        });
      });
    });

  }

  private createDocumentLibrariesJSOM(siteUrl, newSiteColURL): Promise<any> {

    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      var self = this;
      this.setState({
        currentStatus: 'Creating Document Libraries...',
        error: null
      });
      //read the list to get the project structure
      const context: SP.ClientContext = new SP.ClientContext(siteUrl);

      const newContext: SP.ClientContext = new SP.ClientContext(newSiteColURL);
      const newlists: SP.ListCollection = newContext.get_web().get_lists();

      //name of doc library hard coded.
      this.readLibraryConfigurationList(context, "DocumentLibraryStructure").then((items: any) => {
        console.log(items);
        //create doc libraries based on the configuration list

        var loop = function (i) {
          self.createFolderStructure(items[i], newlists, newContext, resolve, self, i, function () {
              if (++i < items.length) {
                  loop(i);
                  
              } else {
                  //act.SharePoint.SharePointAppProgress.completed(true, "Completed");
                    self.setState({
                      currentStatus: 'All folders created',
                      error: null
                    });
                    resolve(true);
              }
          });
      };
      loop(0);
      
      });
    });


  }
  private createFolderStructure(element: any, newlists: SP.ListCollection, newContext: SP.ClientContext, resolve: (itemObjects: any) => void, self: this, index,successFolderNavigation) {
    var docLibCreation: SP.ListCreationInformation;
    docLibCreation = new SP.ListCreationInformation();
    docLibCreation.set_title(element.LibraryName); //list title
    docLibCreation.set_templateType(SP.ListTemplateType.documentLibrary); //document library type
    var newDocLib = newlists.add(docLibCreation);
    newContext.load(newDocLib);
    newContext.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
      //new doc library created now create folders by reading the folders list
      var folder = createFolder(newDocLib, element.FolderStructure);
     
      successFolderNavigation();
      self.setState({
        currentStatus: 'Creating folders -  ' + element.FolderStructure,
        error: null
      });
    }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
      //if doc lib already exists try to create levels -- todo
      const list: SP.List = newContext.get_web().get_lists().getByTitle(element.LibraryName);
      newContext.load(list);
      newContext.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
        var folder = createFolder(list, element.FolderStructure);
        successFolderNavigation();
        self.setState({
          currentStatus: 'Creating folders -  ' + element.FolderStructure,
          error: null
        });
      }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
      });
    });
  }

  // called with the Create Site button is cliecked
  loadForm(event: any) {
    this.setState({ loadForm: true });
  }
  createSiteCollection(formData: IFormData) {
    this.setState({ showCurrentStatus: true, currentStatus: "Provisioning your site" });
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
        self.setState({ showCurrentStatus: true, currentStatus: "Created Site Collection, looking for the site id..." });
        var groupId = data.id;
        setTimeout(function () {
          self.getSiteCollectionIdFromGroupId(self, groupId)
            .then(response => {
              return response.json();
            })
            .then(data => {
              self.setState({ showCurrentStatus: true, currentStatus: "Got the Site Collection ID" });
              var siteCollectionId = data.id;
              self.setState({ currentCreatedSiteUrl: data.webUrl });
              //create document libraries and Folder Structure
              self.createDocumentLibrariesJSOM(self.props.context.pageContext.web.absoluteUrl, data.webUrl).then((data) => {
                self.getListsTitles(self.state.currentCreatedSiteUrl);
                self.setState({ showCurrentStatus: true, currentStatus: "Your brand new team site has been created" });
                Object.assign(document.createElement('a'), { target: '_blank', href: self.state.currentCreatedSiteUrl }).click();

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
    const titles: JSX.Element[] = this.state.listTitles.map((listTitle: string, index: number, listTitles: string[]): JSX.Element => {
      return <li key={index}>{listTitle}</li>;
    });
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
        <br />
        {this.state.loadingLists &&
          <span>Loading lists...</span>}
        {this.state.error &&
          <span>An error has occurred while loading lists: {this.state.error}</span>}
        {this.state.error === null && titles &&
          <ul>
            {titles}
          </ul>}
      </div>
    );
  }
}
