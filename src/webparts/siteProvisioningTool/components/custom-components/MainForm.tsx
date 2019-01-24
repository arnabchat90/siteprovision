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

export interface IMainState {
    init: boolean;
}

export interface IMainProps {
    createSiteCollection: any;
    spContext: any;
}




export default class MainForm extends React.Component<IMainProps, IMainState> {
    constructor(props) {
        super(props);
    }

    private _log(str: string): () => void {
        return (): void => {
            console.log(str);
        };
    }

    private _getPeoplePickerItems(items: any[]) {
        console.log('Items:', items);
      }

    public render(): JSX.Element {
        return (
            <div className={styles.mainForm}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="siteName">Site Name</Label>
                            <TextField id="siteName" />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="siteDescription">Site Description</Label>
                            <TextField multiline rows={4} id="siteDescription" />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="groupEmailAddress">Group Email Address</Label>
                            <TextField id="groupEmailAddress" />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Dropdown
                                placeholder="Select an Option"
                                label="Privacy Settings :"
                                style={{ fontWeight: FontWeights.semibold }}
                                id="privacySettings"
                                ariaLabel="Basic dropdown example"
                                options={[
                                    { key: 'Private', text: 'Private - Only Members Can access this site', title: 'Private - Only Members Can access this site' },
                                    { key: 'Public', text: 'Public - Anyone Can Access This Site' }
                                ]}
                                onFocus={this._log('onFocus called')}
                                onBlur={this._log('onBlur called')}
                            />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Label style={{ fontWeight: FontWeights.semibold }} htmlFor="scOwner">Site Collection Owner</Label>
                           
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <DefaultButton
                                data-automation-id="createSiteCollection"
                                text="Finish"
                                onClick={this.props.createSiteCollection}
                            />
                        </div>
                    </div>
                </div>

            </div>
        );
    }
}