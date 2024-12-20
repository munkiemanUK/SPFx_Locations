import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'LocationsWebPartStrings';
import Locations from './components/Locations';
import { ILocationsProps } from './components/ILocationsProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { FluentProvider, FluentProviderProps, teamsDarkTheme, teamsLightTheme, webLightTheme, webDarkTheme, Theme } from '@fluentui/react-components';

require("bootstrap");

export interface ILocationsWebPartProps {
  description: string;
}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

export default class LocationsWebPart extends BaseClientSideWebPart<ILocationsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme: Theme = webLightTheme;

  public render(): void {
    const element: React.ReactElement<ILocationsProps> = React.createElement(
      Locations,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        webURL: this.context.pageContext.web.absoluteUrl,
        appMode: this._appMode      }
    );

    //wrap the component with the Fluent UI 9 Provider.
    const fluentElement: React.ReactElement<FluentProviderProps> = React.createElement(
      FluentProvider,
      {
        theme: this._appMode === AppMode.Teams || this._appMode === AppMode.TeamsLocal ?
          this._isDarkTheme ? teamsDarkTheme : teamsLightTheme :
          this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal ?
            this._isDarkTheme ? webDarkTheme : this._theme :
            this._isDarkTheme ? webDarkTheme : webLightTheme
      },
      element
    );

    ReactDom.render(fluentElement, this.domElement);
  }

  public async onInit(): Promise<void> {
    await super.onInit();

    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css");
    
    const _l = this.context.isServedFromLocalhost;
    if (!!this.context.sdks.microsoftTeams) {
      const teamsContext = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      switch (teamsContext.app.host.name.toLowerCase()) {
        case 'teams': this._appMode = _l ? AppMode.TeamsLocal : AppMode.Teams; break;
        case 'office': this._appMode = _l ? AppMode.OfficeLocal : AppMode.Office; break;
        case 'outlook': this._appMode = _l ? AppMode.OutlookLocal : AppMode.Outlook; break;
        default: throw new Error('Unknown host');
      }
    } else this._appMode = _l ? AppMode.SharePointLocal : AppMode.SharePoint;

    //const configurableTokenProvider: AadTokenProvider = _AadTokenProvider.configurable as AadTokenProvider;
    /*
    configurableTokenProvider.popupEvent.add(this, (args: IPopupEventArgs) => {
      args.cancel();     // REQUIRED: to cancel the default popup that is called to open
      args.showPopup();  // initiate the popup flow
    });
    */

    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
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
