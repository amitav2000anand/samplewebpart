/*import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
*/

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
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

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _environmentMessage: string = '';
  public async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();

  }
  public render(): void {
    const user = this.context.pageContext.user;

    const element: React.ReactElement<IHelloWorldProps> = React.createElement(HelloWorld, {
      ...this.properties,
      userEmail: user.email,
      userFriendlyName: user.displayName,
      environmentMessage: this._environmentMessage
      /*
      botURL: "https://6a0383e2ebc3ee40bdc9d05198285a.12.environment.api.powerplatform.com/powervirtualagents/botsbyschema/crfb3_selfServiceAgent/directline/token?api-version=2022-03-01-preview",
      botName: "Copilot Self Service Agent",
      buttonLabel: "Ask Copilot",
      botAvatarImage: "", // optional
      botAvatarInitials: "CA", // optional
      greet: true,
      customScope: "api://27eb69cb-53e6-40db-8533-dfa804dca4ac/Sites.Read",
      clientID: "4ffd1a7a-9a30-48a9-bc3d-51060e46591b",
      authority: "https://login.microsoftonline.com/s63fb.onmicrosoft.com",
      userEmail: user.email,
      userFriendlyName: user.displayName,
      environmentMessage: this._environmentMessage
      */
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
                PropertyPaneTextField('botAvatarInitials', { label: 'Bot Avatar Initials' }),
                PropertyPaneTextField('buttonLabel', { label: 'Chat Button Label' })
              ]
            }
          ]
        }
      ]
    };
  }
}
