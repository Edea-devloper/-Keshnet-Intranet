import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'KeshnetPhonebookWebPartStrings';
import KeshetNet from './components/KeshnetPhonebook';
import { IKeshnetPhonebookProps } from './components/IKeshnetPhonebookProps';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import {
  PropertyFieldMultiSelect
} from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";


//  Updated PnP JS imports
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";

export interface ITabConfig {
  tabTitle: string;
  selectedColumns: string[];
  selectedColumnsCard: string[];
}

export interface IKeshnetPhonebookWebPartProps {
  description: string;
  PhoneBookList: string;
  selectedColumns: string[];
  selectedColumnsCard: string[];
  selectordercolumn: string;
  context: any;
  maxRows?: number;
  tabConfigs: ITabConfig[];
  ImageLibrary: string;
  AssetFolderPath: string;
}

export default class KeshnetPhonebookWebPart extends BaseClientSideWebPart<IKeshnetPhonebookWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listColumns: { key: string; text: string }[] = [];
  private _sp: SPFI; // ✅ Added SPFI instance

  public render(): void {
    const element: React.ReactElement<IKeshnetPhonebookProps> = React.createElement(
      KeshetNet,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        PhoneBookList: this.properties.PhoneBookList,
        selectedColumns: this.properties.selectedColumns,
        selectedColumnsCard: this.properties.selectedColumnsCard,
        maxRows: this.properties.maxRows || 10,
        context: this.context,
        tabConfigs: this.properties.tabConfigs,
        listColumns: this._listColumns,
        ImageLibrary: this.properties.ImageLibrary,
        AssetFolderPath: this.properties.AssetFolderPath,
        selectordercolumn: this.properties.selectordercolumn
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {

    // Initialize PnP 
    this._sp = spfi().using(SPFx(this.context));

    // ✅ Preload columns if a list is already selected
    if (this.properties.PhoneBookList) {
      await this._loadColumns(this.properties.PhoneBookList);
    }

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

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === "PhoneBookList" && newValue) {
      // Reset selected columns when list changes
      this.properties.selectedColumns = [];
      this.properties.selectedColumnsCard = [];

      await this._loadColumns(newValue);

      // Refresh property pane so new column options show up
      this.context.propertyPane.refresh();

      // Re-render web part after reset
      this.render();
    }
  }


  // Updated to use this._sp instead of old sp
  // Updated to include Calculated columns (FieldTypeKind 17)
  private async _loadColumns(listId: string): Promise<void> {
    const fields = await this._sp.web.lists.getById(listId).fields
      .select("InternalName,Title,Hidden,ReadOnlyField,FieldTypeKind") // include FieldTypeKind
      .filter("Hidden eq false")(); // only exclude hidden

    this._listColumns = fields
      .filter((f: any) => {
        // Allow all non-readonly OR explicitly allow calculated (FieldTypeKind 17)
        return !f.ReadOnlyField || f.FieldTypeKind === 17;
      })
      .map((f: { InternalName: string; Title: string }) => ({
        key: f.InternalName,
        text: f.Title
      }));
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
                }),
                PropertyFieldListPicker("PhoneBookList", {
                  label: "Select PhoneBook List",
                  selectedList: this.properties.PhoneBookList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "PhoneBookListPicker",
                }),
                PropertyFieldMultiSelect("selectedColumns", {
                  key: "multiSelectColumns",
                  label: "Select Columns",
                  options: this._listColumns,
                  selectedKeys: this.properties.selectedColumns || [],
                  disabled: this._listColumns.length === 0,
                  //properties: this.properties
                }),
                PropertyFieldMultiSelect("selectedColumnsCard", {
                  key: "multiSelectColumns",
                  label: "Select Columns for Card",
                  options: this._listColumns,
                  selectedKeys: this.properties.selectedColumnsCard || [],
                  disabled: this._listColumns.length === 0,
                  //properties: this.properties
                }),
                PropertyPaneDropdown('selectordercolumn', {
                  label: 'Select column for ordering',
                  options: this._listColumns,
                  selectedKey: this.properties.selectordercolumn || '',
                  disabled: this._listColumns.length === 0,
                }),
                PropertyPaneTextField('maxRows', {
                  label: 'Max Rows to Display',
                  description: 'Enter a number',
                  //type: 'number',
                  value: this.properties.maxRows !== undefined ? String(this.properties.maxRows) : undefined, // default to 5
                }),
                PropertyFieldListPicker('ImageLibrary', {
                  label: 'Select Image Library',
                  selectedList: this.properties.ImageLibrary,
                  includeHidden: false,
                  baseTemplate: 109, // <-- filters only document libraries
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: 'libraryPickerFieldId'
                }),
                PropertyPaneTextField('AssetFolderPath', {
                  label: 'Enter Asset FolderPath',
                  description: 'Provide the Asset FolderPath',
                  value: this.properties.AssetFolderPath
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
