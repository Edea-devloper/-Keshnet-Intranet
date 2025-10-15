import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneTextField,
  type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'KeshnetBirthdayWebPartStrings';
import Birthday from './components/KeshnetBirthday';
import { IKeshnetBirthdayProps } from './components/IKeshnetBirthdayProps';
import { getSP } from './Utility/getSP';
import { getListItemsFormainList } from './Utility/utils';
import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';

export interface IBirthdayWebPartProps {
  selectedList: string | string[];
  description: string;
  fullname: string;
  department: string;
  company: string;
  ImageLibrary: string;
  Duration: string;
}

export default class BirthdayWebPart extends BaseClientSideWebPart<IBirthdayWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listData: any[] = []; // store fetched list data

  public render(): void {
    const element: React.ReactElement<IKeshnetBirthdayProps> = React.createElement(
      Birthday,
      {
        context: this.context,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        selectedList: this.properties.selectedList,
        listData: this._listData,
        fullname: this.properties.fullname,
        department: this.properties.department,
        company: this.properties.company,
        ImageLibrary: this.properties.ImageLibrary,
        Duration: this.properties.Duration
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();
    getSP(this.context);

    if (this.properties.selectedList) {
      try {
        this._listData = await getListItemsFormainList({ title: this.properties.selectedList });
        console.log("Fetched list data:", this._listData);
      } catch (error) {
        console.error("Error fetching list data:", error);
      }
    }

    return;
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
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
    const { semanticColors } = currentTheme;

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
          groups: [
            {
              groupFields: [
                PropertyFieldListPicker('selectedList', {
                  label: 'Select a SharePoint list',
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldColumnPicker('fullname', {
                  label: 'Select a FullName column',
                  context: this.context,
                  selectedColumn: this.properties.fullname,
                  listId: this.properties.selectedList ? this.properties.selectedList.toString() : "",
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'fullnameColumnPicker',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('department', {
                  label: 'Select a Department column',
                  context: this.context,
                  selectedColumn: this.properties.department,
                  listId: this.properties.selectedList ? this.properties.selectedList.toString() : "",
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'departmentColumnPicker',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldColumnPicker('company', {
                  label: 'Select a Company column',
                  context: this.context,
                  selectedColumn: this.properties.company,
                  listId: this.properties.selectedList ? this.properties.selectedList.toString() : "",
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'companyColumnPicker',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
                }),
                PropertyFieldListPicker('ImageLibrary', {
                  label: 'Select Image Library',
                  selectedList: this.properties.ImageLibrary,
                  includeHidden: false,
                  baseTemplate: 109,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: 'libraryPickerFieldId'
                }),
                PropertyPaneTextField('Duration', {
                  label: 'Enter seconds to change the content'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}