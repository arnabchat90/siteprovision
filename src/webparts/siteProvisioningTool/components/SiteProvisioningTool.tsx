import * as React from 'react';
import styles from './SiteProvisioningTool.module.scss';
import { ISiteProvisioningToolProps } from './ISiteProvisioningToolProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext, IWebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('taxonomy');

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
  allPermissionLevels: any[];
  allTerms: any[];
  formData: IFormData;
  messageBarType : MessageBarType;

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
import SPTaxonomyService from '../SPTaxonomyService';

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
      error: '',
      allPermissionLevels: [],
      allTerms: null,
      formData: null,
      messageBarType : MessageBarType.success
    };
    this.loadForm = this.loadForm.bind(this);
    this.createSiteCollection = this.createSiteCollection.bind(this);
    this.getListsTitles = this.getListsTitles.bind(this);
    this.createDocumentLibrariesJSOM = this.createDocumentLibrariesJSOM.bind(this);
    this.readLibraryConfigurationList = this.readLibraryConfigurationList.bind(this);
    this.createFolderStructure = this.createFolderStructure.bind(this);
    this.createCustomPermissionLevels = this.createCustomPermissionLevels.bind(this);
    this.readPermissionList = this.readPermissionList.bind(this);
    this.createSharePointGroups = this.createSharePointGroups.bind(this);
    this.createSharePointGroupinSP = this.createSharePointGroupinSP.bind(this);
    this.readFromTermSet = this.readFromTermSet.bind(this);
    this.addItemToProjectsList = this.addItemToProjectsList.bind(this);
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
  private async readFromTermSet(context: IWebPartContext, termSetname: string) {
    var self = this;
    let svc: SPTaxonomyService = new SPTaxonomyService(context);
    let result = await svc.getTermsFromTermSet(termSetname);
    var allTermsOfPractice = result.map((term) => ({
      key: term.Id,
      text: term.Name
    }));
    self.setState({ allTerms: allTermsOfPractice });
  }

  private awaitSPExecuteQuery(context: SP.ClientContext): Promise<any> {
    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      context.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
        resolve(true);
      }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
        reject(false);
      });
    });

  }

  private async addItemToProjectsList(context: SP.ClientContext, listName: string) {
    var self = this;
    var oList = context.get_web().get_lists().getByTitle(listName);
    var itemCreateInfo = new SP.ListItemCreationInformation();
    var oListItem = oList.addItem(itemCreateInfo);
    var field = oList.get_fields().getByInternalNameOrTitle('Practice');
    var txField = context.castTo(field, SP.Taxonomy.TaxonomyField);
    var website = context.get_web();
    var pmUser = website.ensureUser(self.state.formData.pm[0].id);
    var execUser = website.ensureUser(self.state.formData.executive[0].id);
    var spUser = website.ensureUser(self.state.formData.sp[0].id);
    context.load(website);
    context.load(pmUser);
    context.load(execUser);
    context.load(spUser);
    context.load(txField);
    await self.awaitSPExecuteQuery(context);
    //1. Prepare TaxonomyFieldValue
    var termValue = new SP.Taxonomy.TaxonomyFieldValue();
    termValue.set_label(self.state.formData.practice.text);
    termValue.set_termGuid(self.state.formData.practice.key);
    termValue.set_wssId(-1);
    oListItem.set_item('Client_x0020_Name', self.state.formData.siteName);
    var pmVal = new SP.FieldUserValue();
    pmVal.set_lookupId(pmUser.get_id());   //specify User Id 
    oListItem.set_item('PM', pmVal);
    var execVal = new SP.FieldUserValue();
    execVal.set_lookupId(execUser.get_id());   //specify User Id 
    oListItem.set_item('Executive', execVal);
    var spVal = new SP.FieldUserValue();
    spVal.set_lookupId(spUser.get_id());   //specify User Id 
    oListItem.set_item('Sales_x0020_Person', spVal);
    txField.setFieldValueByValue(oListItem, termValue);
    oListItem.update();
    context.load(oListItem);
    await self.awaitSPExecuteQuery(context);

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

  private readPermissionList(clientContext: SP.ClientContext, listName: string): Promise<any> {
    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      var self = this;
      var oList = clientContext.get_web().get_lists().getByTitle(listName);
      var query = SP.CamlQuery.createAllItemsQuery();
      var collListItem = oList.getItems(query);
      clientContext.load(collListItem);
      clientContext.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
        const listEnumerator: IEnumerator<SP.ListItem> = collListItem.getEnumerator();
        let itemObjects = [];

        while (listEnumerator.moveNext()) {
          const listItem: SP.ListItem = listEnumerator.get_current();

          itemObjects.push({
            SPGroupName: listItem.get_item('Title'),
            CustomPermissionName: listItem.get_item('PermissionLevel')
          });

        }
        self.setState({
          currentStatus: 'Reading All Permissions From Mapping List',
          error: null,
          messageBarType : MessageBarType.info
        });
        resolve(itemObjects);

      }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
        self.setState({
          showCurrentStatus: true, currentStatus: "There is some error while provisioning your site" + args.get_message(),
          messageBarType : MessageBarType.error
        });
        reject(args.get_message());
      });
    });

  }

  private createDocumentLibrariesJSOM(siteUrl, newSiteColURL): Promise<any> {

    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      var self = this;
      this.setState({
        currentStatus: 'Creating Document Libraries...',
        error: null,
        messageBarType : MessageBarType.info
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
                error: null,
                messageBarType : MessageBarType.info
              });
              resolve(true);
            }
          });
        };
        loop(0);

      });
    });


  }
  private createFolderStructure(element: any, newlists: SP.ListCollection, newContext: SP.ClientContext, resolve: (itemObjects: any) => void, self: this, index, successFolderNavigation) {
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
        error: null,
        messageBarType : MessageBarType.info
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
          error: null,
          messageBarType : MessageBarType.info
        });
      }, (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
        self.setState({
          showCurrentStatus: true, currentStatus: "There is some error while provisioning  your site",
          messageBarType : MessageBarType.error
        });
      });
    });
  }

  createSharePointGroups(newSiteUrl: string, spGroupPermissionMapping: any[]): Promise<any> {
    var self = this;
    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      const newContext: SP.ClientContext = new SP.ClientContext(newSiteUrl);

      var loop = function (i) {
        self.createSharePointGroupinSP(newContext, spGroupPermissionMapping[i].SPGroupName, spGroupPermissionMapping[i].SPGroupName, spGroupPermissionMapping[i].CustomPermissionName, function () {
          if (++i < spGroupPermissionMapping.length) {
            self.setState({
              currentStatus: 'Creating SharePoint Group - ' + spGroupPermissionMapping[i].SPGroupName,
              error: null,
              messageBarType : MessageBarType.info
            });
            loop(i);

          } else {
            //act.SharePoint.SharePointAppProgress.completed(true, "Completed");
            self.setState({
              currentStatus: 'All SharePoint Groups Created',
              error: null,
              messageBarType : MessageBarType.info
            });
            resolve(true);
          }
        }, function (error) {
          reject(error);
          self.setState({
            showCurrentStatus: true, currentStatus: "There is some error - " + error.message,
            messageBarType : MessageBarType.error
          });
        });
      };
      loop(0);
    });
  }

  createCustomPermissionLevels(newSiteUrl: string): Promise<any> {
    var self = this;
    return new Promise<any>((resolve: (itemObjects: any) => void, reject: (error: any) => void): void => {
      const newContext: SP.ClientContext = new SP.ClientContext(newSiteUrl);
      var customPermissionLevelNames = ['FULL ACCESS', 'Edit With Delete', 'Edit With No Delete', 'Read Only'];
      var dsReadPermissions = [];
      customPermissionLevelNames.forEach((item) => {
        dsReadPermissions.push({ levelName: item, permissions: this.createPermissionSet(item) });
      });
      self.setState({
        allPermissionLevels: dsReadPermissions
      });
      var loop = function (i) {
        self.createCustomPermission(newContext, dsReadPermissions[i].levelName, dsReadPermissions[i].levelName, dsReadPermissions[i].permissions, function () {
          if (++i < dsReadPermissions.length) {
            self.setState({
              currentStatus: 'Creating Custom Permission - ' + dsReadPermissions[i].levelName,
              error: null,
              messageBarType : MessageBarType.info
            });
            loop(i);

          } else {
            //act.SharePoint.SharePointAppProgress.completed(true, "Completed");
            self.setState({
              currentStatus: 'All Custom Permissions Created',
              error: null,
              messageBarType : MessageBarType.info
            });
            resolve(true);
          }
        }, function (error) {
          reject(error);
        });
      };
      loop(0);
    });
  }
  createCustomPermission(context, name, desc, permissions, success, fail) {
    // Create a new role definition.  
    var roleDefinitionCreationInfo = new SP.RoleDefinitionCreationInformation();
    roleDefinitionCreationInfo.set_name(name);
    roleDefinitionCreationInfo.set_description(desc);
    roleDefinitionCreationInfo.set_basePermissions(permissions);
    var roleDefinition = context.get_site().get_rootWeb().get_roleDefinitions().add(roleDefinitionCreationInfo);
    context.executeQueryAsync(success, fail);
  }

  createSharePointGroupinSP(context, name, desc, roleName, success, fail) {
    var web = context.get_web();
    //Get all groups in site  
    var groupCollection = web.get_siteGroups();
    // Create Group information for Group  
    var newGRP = new SP.GroupCreationInformation();
    newGRP.set_title(name);
    newGRP.set_description(desc);
    // if (!web.get_hasUniqueRoleAssignments()) {
    //   web.breakRoleInheritance(true, false);
    // }
    //add group to site gorup collection  
    var newCreateGroup = groupCollection.add(newGRP);
    //Role Definition   
    var rolDef = web.get_roleDefinitions().getByName(roleName);
    var rolDefColl = SP.RoleDefinitionBindingCollection.newObject(context);
    rolDefColl.add(rolDef);

    // Get the RoleAssignmentCollection for the target web.  
    var roleAssignments = web.get_roleAssignments();
    // assign the group to the new RoleDefinitionBindingCollection.  
    roleAssignments.add(newCreateGroup, rolDefColl);
    //Set group properties  
    newCreateGroup.set_allowMembersEditMembership(true);
    newCreateGroup.set_onlyAllowMembersViewMembership(false);
    newCreateGroup.update();
    context.executeQueryAsync(success, fail);
  }

  createPermissionSet(perm: string) {
    var permissions = new SP.BasePermissions();
    switch (perm) {
      case "FULL ACCESS":
        permissions.set(SP.PermissionKind.manageLists);
        permissions.set(SP.PermissionKind.cancelCheckout);
        permissions.set(SP.PermissionKind.addListItems);
        permissions.set(SP.PermissionKind.editListItems);
        permissions.set(SP.PermissionKind.deleteListItems);
        permissions.set(SP.PermissionKind.viewListItems);
        permissions.set(SP.PermissionKind.approveItems);
        permissions.set(SP.PermissionKind.openItems);
        permissions.set(SP.PermissionKind.viewVersions);
        permissions.set(SP.PermissionKind.deleteVersions);
        permissions.set(SP.PermissionKind.createAlerts);
        permissions.set(SP.PermissionKind.viewPages);
        permissions.set(SP.PermissionKind.viewFormPages);
        permissions.set(SP.PermissionKind.browseDirectories);
        permissions.set(SP.PermissionKind.createSSCSite);
        permissions.set(SP.PermissionKind.addAndCustomizePages);
        permissions.set(SP.PermissionKind.browseUserInfo);
        permissions.set(SP.PermissionKind.useRemoteAPIs);
        permissions.set(SP.PermissionKind.useClientIntegration);
        permissions.set(SP.PermissionKind.open);
        permissions.set(SP.PermissionKind.editMyUserInfo);
        permissions.set(SP.PermissionKind.managePersonalViews);
        permissions.set(SP.PermissionKind.addDelPrivateWebParts);
        permissions.set(SP.PermissionKind.updatePersonalWebParts);
        break;
      case "Edit With Delete":
        // permissions.set(SP.PermissionKind.manageLists);
        // permissions.set(SP.PermissionKind.cancelCheckout);
        permissions.set(SP.PermissionKind.addListItems);
        permissions.set(SP.PermissionKind.editListItems);
        permissions.set(SP.PermissionKind.deleteListItems);
        permissions.set(SP.PermissionKind.viewListItems);
        //permissions.set(SP.PermissionKind.approveItems);
        permissions.set(SP.PermissionKind.openItems);
        permissions.set(SP.PermissionKind.viewVersions);
        permissions.set(SP.PermissionKind.createAlerts);
        permissions.set(SP.PermissionKind.viewPages);
        permissions.set(SP.PermissionKind.viewFormPages);
        permissions.set(SP.PermissionKind.browseDirectories);
        permissions.set(SP.PermissionKind.createSSCSite);
        permissions.set(SP.PermissionKind.addAndCustomizePages);
        permissions.set(SP.PermissionKind.browseUserInfo);
        permissions.set(SP.PermissionKind.useRemoteAPIs);
        permissions.set(SP.PermissionKind.useClientIntegration);
        permissions.set(SP.PermissionKind.open);
        permissions.set(SP.PermissionKind.editMyUserInfo);
        permissions.set(SP.PermissionKind.managePersonalViews);
        permissions.set(SP.PermissionKind.addDelPrivateWebParts);
        permissions.set(SP.PermissionKind.updatePersonalWebParts);
        break;
      case "Edit With No Delete":
        // permissions.set(SP.PermissionKind.manageLists);
        // permissions.set(SP.PermissionKind.cancelCheckout);
        permissions.set(SP.PermissionKind.addListItems);
        permissions.set(SP.PermissionKind.editListItems);
        //permissions.set(SP.PermissionKind.deleteListItems);
        permissions.set(SP.PermissionKind.viewListItems);
        //permissions.set(SP.PermissionKind.approveItems);
        permissions.set(SP.PermissionKind.openItems);
        permissions.set(SP.PermissionKind.viewVersions);
        permissions.set(SP.PermissionKind.createAlerts);
        permissions.set(SP.PermissionKind.viewPages);
        permissions.set(SP.PermissionKind.viewFormPages);
        permissions.set(SP.PermissionKind.browseDirectories);
        permissions.set(SP.PermissionKind.createSSCSite);
        permissions.set(SP.PermissionKind.addAndCustomizePages);
        permissions.set(SP.PermissionKind.browseUserInfo);
        permissions.set(SP.PermissionKind.useRemoteAPIs);
        permissions.set(SP.PermissionKind.useClientIntegration);
        permissions.set(SP.PermissionKind.open);
        permissions.set(SP.PermissionKind.editMyUserInfo);
        permissions.set(SP.PermissionKind.managePersonalViews);
        permissions.set(SP.PermissionKind.addDelPrivateWebParts);
        permissions.set(SP.PermissionKind.updatePersonalWebParts);
        break;
      case "Read Only":
        // permissions.set(SP.PermissionKind.manageLists);
        // permissions.set(SP.PermissionKind.cancelCheckout);
        //permissions.set(SP.PermissionKind.addListItems);
        //permissions.set(SP.PermissionKind.editListItems);
        //permissions.set(SP.PermissionKind.deleteListItems);
        permissions.set(SP.PermissionKind.viewListItems);
        //permissions.set(SP.PermissionKind.approveItems);
        permissions.set(SP.PermissionKind.openItems);
        permissions.set(SP.PermissionKind.viewVersions);
        permissions.set(SP.PermissionKind.createAlerts);
        permissions.set(SP.PermissionKind.viewPages);
        permissions.set(SP.PermissionKind.viewFormPages);
        //permissions.set(SP.PermissionKind.browseDirectories);
        permissions.set(SP.PermissionKind.createSSCSite);
        //permissions.set(SP.PermissionKind.addAndCustomizePages);
        permissions.set(SP.PermissionKind.browseUserInfo);
        permissions.set(SP.PermissionKind.useRemoteAPIs);
        permissions.set(SP.PermissionKind.useClientIntegration);
        permissions.set(SP.PermissionKind.open);
        //permissions.set(SP.PermissionKind.editMyUserInfo);
        //permissions.set(SP.PermissionKind.managePersonalViews);
        //permissions.set(SP.PermissionKind.addDelPrivateWebParts);
        //permissions.set(SP.PermissionKind.updatePersonalWebParts);
        break;

      default:
        break;
    }
    return permissions;
  }

  // called with the Create Site button is cliecked
  loadForm(event: any) {
    this.setState({ loadForm: true });
  }
  createSiteCollection(formData: IFormData) {
    
    
    var self = this;
    self.setState({ formData: formData });
    if(formData.siteName == null || formData.siteName == undefined || formData.siteName == '') {
      self.setState({ showCurrentStatus: true, currentStatus: "Site Name can not be empty",
        messageBarType : MessageBarType.error });
      return;
    }

    if(formData.siteDescription == null || formData.siteDescription == undefined || formData.siteDescription == '') {
      self.setState({ showCurrentStatus: true, currentStatus: "Site Description can not be empty",
        messageBarType : MessageBarType.error });
      return;
    }
    
    self.setState({ showCurrentStatus: true, currentStatus: "Provisioning your site",
    messageBarType : MessageBarType.info });
    const siteCreationBody: string = JSON.stringify(
      {
        description: formData.siteDescription,
        displayName: formData.siteName,
        groupTypes: [
          "Unified"
        ],
        mailEnabled: false,
        mailNickname: formData.groupEmailAddress,
        securityEnabled: false,
        visibility: formData.privacyOptions.key
      });
    //create the site collection using graph api
    self.createNewSiteCollectionUsingGraph(self, siteCreationBody)
      .then(response => {
        return response.json();
      }).then(data => {
        console.log(data);
        //get root site collection id from group id
        self.setState({ showCurrentStatus: true, currentStatus: "Created Site Collection, looking for the site id...",
        messageBarType : MessageBarType.info });
        var groupId = data.id;
        setTimeout(function () {
          self.getSiteCollectionIdFromGroupId(self, groupId)
            .then(response => {
              return response.json();
            })
            .then(data => {
              self.setState({ showCurrentStatus: true, currentStatus: "Got the Site Collection ID",
              messageBarType : MessageBarType.info });
              console.log(data);
              if (data.error !== null && data.error !== undefined) {
                self.setState({
                  showCurrentStatus: true, currentStatus: "There is some error - " + data.error.message,
                  messageBarType : MessageBarType.error
                });
                return;
              }
              var siteCollectionId = data.id;
              self.setState({ currentCreatedSiteUrl: data.webUrl });
              //create document libraries and Folder Structure
              self.createDocumentLibrariesJSOM(self.props.context.pageContext.web.absoluteUrl, data.webUrl).then((data) => {
                //self.getListsTitles(self.state.currentCreatedSiteUrl);
                self.createCustomPermissionLevels(self.state.currentCreatedSiteUrl).then(() => {
                  //read group to permission mappings from list, then add groups and assign permissions
                  const context: SP.ClientContext = new SP.ClientContext(self.props.context.pageContext.web.absoluteUrl);
                  self.readPermissionList(context, "PermissionAssignment").then((groupPermissionMappings) => {
                    //create the actual groups in sharepoint with the permissions
                    self.createSharePointGroups(self.state.currentCreatedSiteUrl, groupPermissionMappings).then(() => {
                      //add item to the projects list and then show success
                      self.addItemToProjectsList(context, 'Projects');
                      self.setState({ showCurrentStatus: true, currentStatus: "Your brand new team site has been created",
                      messageBarType : MessageBarType.success });
                      Object.assign(document.createElement('a'), { target: '_blank', href: self.state.currentCreatedSiteUrl }).click();
                    });
                  });

                });

              });
            });
        }, 10000);
      }, (error) => {
        self.setState({
          showCurrentStatus: true, currentStatus: "There is some error while creating - " + error.message,
          messageBarType : MessageBarType.error
        });
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
    //read term store and get Practice values
    this.readFromTermSet(this.props.context, 'Practice');

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

        </div> : <MainForm createSiteCollection={this.createSiteCollection} spContext={this.props.context} practiceTerms={this.state.allTerms} currentStatus={this.state.currentStatus} messageBarType={this.state.messageBarType} showCurrentStatus={this.state.showCurrentStatus} />}
        <br />
        {/* {this.state.loadingLists &&
          <span>Loading lists...</span>}
        {this.state.error &&
          <span>An error has occurred while loading lists: {this.state.error}</span>}
        {this.state.error === null && titles &&
          <ul>
            {titles}
          </ul>} */}
      </div>
    );
  }
}
