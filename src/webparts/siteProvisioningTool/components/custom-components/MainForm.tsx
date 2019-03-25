import * as React from 'react';
import styles from './MainForm.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { FontWeights } from '@uifabric/styling/lib';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { BaseComponent } from 'office-ui-fabric-react/lib/Utilities';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
    SPHttpClient,
    SPHttpClientResponse
} from '@microsoft/sp-http';

import { MessageBarButton } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export interface IMainState {
    init: boolean;
    formData: IFormData;
}

export interface IFormData {
    siteName: string;
    siteDescription: string;
    groupEmailAddress: string;
    siteOwners: Array<any>;
    pm: Array<any>;
    executive: Array<any>;
    sp: Array<any>;
    privacyOptions: any;
    practice: any;
}

export interface IMainProps {
    createSiteCollection: any;
    spContext: any;
    currentStatus: string;
    showCurrentStatus: boolean;
    practiceTerms: any[];
    messageBarType : MessageBarType;
}

export default class MainForm extends React.Component<IMainProps, IMainState> {
    constructor(props) {
        super(props);
        this.state = {
            init: false,
            formData: {
                siteName: "",
                siteDescription: "",
                groupEmailAddress: "",
                siteOwners: [],
                pm: [],
                executive: [],
                sp: [],
                privacyOptions: {},
                practice: {}
            }
        }

        this._handleOnChange = this._handleOnChange.bind(this);
        this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
        this._getPeoplePickerItemsPM = this._getPeoplePickerItemsPM.bind(this);
        this._getPeoplePickerItemsExecutive = this._getPeoplePickerItemsExecutive.bind(this);
        this._getPeoplePickerItemsSP = this._getPeoplePickerItemsSP.bind(this);
    }

    _handleOnChange(val: any, fieldName: string) {
        var tempFormData = { ...this.state.formData };
        switch (fieldName) {
            case "siteName":
                tempFormData.siteName = val;
                tempFormData.groupEmailAddress = val.replace(/ /g, '');
                break;
            case "siteDescription":
                tempFormData.siteDescription = val;
                break;
            case "groupEmailAddress":
                tempFormData.groupEmailAddress = val;
                break;
            case "privacyOptions":
                tempFormData.privacyOptions = val;
                break;
            case "practice":
                tempFormData.practice = val;
                break;
        }
        this.setState({ formData: tempFormData });
    }

    private _log(str: string): () => void {
        return (): void => {
            console.log(str);
        };
    }

    private _getPeoplePickerItems(items: any[]) {
        var tempFormData = { ...this.state.formData };
        tempFormData.siteOwners = items;
        this.setState({ formData: tempFormData });
    }
    private _getPeoplePickerItemsPM(items: any[]) {
        var tempFormData = { ...this.state.formData };
        tempFormData.pm = items;
        this.setState({ formData: tempFormData });
    }
    private _getPeoplePickerItemsExecutive(items: any[]) {
        var tempFormData = { ...this.state.formData };
        tempFormData.executive = items;
        this.setState({ formData: tempFormData });
    }
    private _getPeoplePickerItemsSP(items: any[]) {
        var tempFormData = { ...this.state.formData };
        tempFormData.sp = items;
        this.setState({ formData: tempFormData });
    }

    public render(): JSX.Element {
        return (
            <div className={styles.mainForm}>
                <div className={styles.container}>
                    {this.props.showCurrentStatus == true ? <div className={styles.row}>
                        <div className={styles.column}>
                            <MessageBar
                                messageBarType={this.props.messageBarType}
                                isMultiline={true}
                            >
                                {this.props.currentStatus}
                            </MessageBar>
                        </div>
                    </div> : ""}

                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="siteName">Site Name *</Label>
                            <TextField id="siteName" value={this.state.formData.siteName} onChanged={(newValue) => { this._handleOnChange(newValue, "siteName") }} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="siteDescription">Site Description *</Label>
                            <TextField multiline rows={4} id="siteDescription" value={this.state.formData.siteDescription} onChanged={(newValue) => { this._handleOnChange(newValue, "siteDescription") }} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="groupEmailAddress">Group Email Address</Label>
                            <TextField id="groupEmailAddress" value={this.state.formData.groupEmailAddress} onChanged={(newValue) => { this._handleOnChange(newValue, "groupEmailAddress") }} />
                        </div>
                    </div>

                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Dropdown
                                placeholder="Select an Option"
                                label="Practice"
                                style={{ fontWeight: FontWeights.semibold }}
                                id="practice"
                                ariaLabel="Basic dropdown example"
                                options={this.props.practiceTerms}
                                onFocus={this._log('onFocus called')}
                                onBlur={this._log('onBlur called')}
                                onChanged={(newValue) => { this._handleOnChange(newValue, "practice") }}
                            />
                        </div>
                    </div>

                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Dropdown
                                placeholder="Select an Option"
                                label="Privacy Settings"
                                style={{ fontWeight: FontWeights.semibold }}
                                id="privacySettings"
                                ariaLabel="Basic dropdown example"
                                options={[
                                    { key: 'Private', text: 'Private - Only Members Can access this site', title: 'Private - Only Members Can access this site' },
                                    { key: 'Public', text: 'Public - Anyone Can Access This Site' }
                                ]}
                                onFocus={this._log('onFocus called')}
                                onBlur={this._log('onBlur called')}
                                onChanged={(newValue) => { this._handleOnChange(newValue, "privacyOptions") }}
                            />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <PeoplePicker
                                context={this.props.spContext}
                                titleText="Project Manager"
                                personSelectionLimit={3}
                                showtooltip={true}
                                isRequired={true}
                                selectedItems={this._getPeoplePickerItemsPM}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={1000} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <PeoplePicker
                                context={this.props.spContext}
                                titleText="Executive"
                                personSelectionLimit={3}
                                showtooltip={true}
                                isRequired={true}
                                selectedItems={this._getPeoplePickerItemsExecutive}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={500} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <PeoplePicker
                                context={this.props.spContext}
                                titleText="Sales Person"
                                personSelectionLimit={3}
                                showtooltip={true}
                                isRequired={true}
                                selectedItems={this._getPeoplePickerItemsSP}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={500} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <PeoplePicker
                                context={this.props.spContext}
                                titleText="Site Collection Owner"
                                personSelectionLimit={3}
                                showtooltip={true}
                                isRequired={true}
                                selectedItems={this._getPeoplePickerItems}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={500} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <PrimaryButton
                                data-automation-id="createSiteCollection"
                                text="Finish"
                                onClick={this.props.createSiteCollection.bind(this, this.state.formData)}
                            />
                        </div>
                    </div>
                </div>

            </div>
        );
    }
}