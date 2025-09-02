import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ServiceDeskChatWebPartStrings';
import ServiceDeskChat from './components/ServiceDeskChat';
import { IServiceDeskChatProps } from './components/IServiceDeskChatProps';

export interface IServiceDeskChatWebPartProps {
  botURL: string;
  botName?: string;
  buttonLabel?: string;
  botAvatarImage?: string;
  botAvatarInitials?: string;
  greet?: boolean;
  customScope: string;
  clientID: string;
  authority: string;
}

export default class ServiceDeskChatWebPart extends BaseClientSideWebPart<IServiceDeskChatWebPartProps> {
  private _environmentMessage: string = '';
  public async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();

  }
  public render(): void {
    const user = this.context.pageContext.user;

    const element: React.ReactElement<IServiceDeskChatProps> = React.createElement(ServiceDeskChat, {
      ...this.properties,
      userEmail: user.email,
      userFriendlyName: user.displayName,
      environmentMessage: this._environmentMessage
    });

    ReactDom.render(element, this.domElement);
  }
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          switch (context.app.host.name) {
            case 'Office':
              return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
            case 'Outlook':
              return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
            case 'Teams':
            case 'TeamsModern':
              return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
            default:
              return strings.UnknownEnvironment;
          }
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }
  public onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('botName', { label: 'Bot Name' }),
                PropertyPaneTextField('botURL', { label: 'Bot URL' }),
                PropertyPaneTextField('clientID', { label: 'Client ID' }),
                PropertyPaneTextField('authority', { label: 'Authority' }),
                PropertyPaneTextField('customScope', { label: 'Custom Scope' }),
                PropertyPaneToggle('greet', { label: 'Greet on Start', onText: 'Yes', offText: 'No' }),
                PropertyPaneTextField('botAvatarImage', { label: 'Bot Avatar Image URL' }),
                PropertyPaneTextField('botAvatarInitials', { label: 'Bot Avatar Initials' })
              ]
            }
          ]
        }
      ]
    };
  }
}
