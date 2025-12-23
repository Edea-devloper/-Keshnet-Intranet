import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "KeshnetQuicklinksWebPartStrings";
import QuickLinks from "./components/KeshnetQuicklinks";
import { IKeshnetQuicklinksProps } from "./components/IKeshnetQuicklinksProps";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import { getSP } from "./Utility/getSP";

export interface IQuickLinksWebPartProps {
  selectedList: string | string[];
  description: string;
  PhoneBookListForQuickLink: string;
  AssetFolderPath: string;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private _listData: any[] = [];
  private _currentUser: any = null; // store current user info
  private _ListDataEMPFirstName: any[] = [];

  public render(): void {
    const element: React.ReactElement<IKeshnetQuicklinksProps> =
      React.createElement(QuickLinks, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        selectedList: this.properties.selectedList,
        listData: this._listData,
        currentUser: this._currentUser,
        PhoneBookListForQuickLink: this.properties.PhoneBookListForQuickLink,
        _ListDataEMPFirstName: this._ListDataEMPFirstName,
        AssetFolderPath: this.properties.AssetFolderPath,
        context: this.context
      });

    ReactDom.render(element, this.domElement);
  }

   protected async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();
    getSP(this.context); // initialize PnP JS
    return;
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyFieldListPicker("selectedList", {
                  label: "Select a SharePoint list For Quick Links",
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("PhoneBookListForQuickLink", {
                  label: "Select PhoneBook List",
                  selectedList: this.properties.PhoneBookListForQuickLink,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "PhoneBookListForQuickLink",
                }),
                PropertyPaneTextField("AssetFolderPath", {
                  label: "Enter Asset FolderPath",
                  description: "Provide the Asset FolderPath",
                  value: this.properties.AssetFolderPath,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
