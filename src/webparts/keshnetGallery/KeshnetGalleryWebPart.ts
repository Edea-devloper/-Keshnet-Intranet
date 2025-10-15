import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'KeshnetGalleryWebPartStrings';
import KeshnetGallery from './components/KeshnetGallery';
import { IKeshnetGalleryProps } from './components/IKeshnetGalleryProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { initializeSP } from './services/SPService';

import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IPropertyControlsTestWebPartProps {
  datetime: IDateTimeFieldValue;
}

export interface IPropertyControlsTestWebPartProps {
  lists: string | string[]; // Stores the list ID(s)
}

export interface IKeshnetGalleryWebPartProps {
  description: string;
  _GalleryLists: string;
  ImageLibrary: string;
  TemplateList: string;
  EmployeeList: string;
  AssetFolderPath: string;
}

export default class KeshnetGalleryWebPart extends BaseClientSideWebPart<IKeshnetGalleryWebPartProps> {



  public render(): void {
    SPComponentLoader.loadCss("https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;600;700&display=swap");
    const element: React.ReactElement<IKeshnetGalleryProps> = React.createElement(
      KeshnetGallery,
      {
        description: this.properties.description,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        _GalleryLists: this.properties._GalleryLists,
        context: this.context,
        ImageLibrary: this.properties.ImageLibrary,
        TemplateList: this.properties.TemplateList,
        EmployeeList: this.properties.EmployeeList,
        AssetFolderPath: this.properties.AssetFolderPath
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await initializeSP(this.context);
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
                }),
                PropertyFieldListPicker('_GalleryLists', {
                  label: 'Select Gallery list',
                  selectedList: this.properties._GalleryLists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
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
                PropertyFieldListPicker('TemplateList', {
                  label: 'Select Gallery Template list',
                  selectedList: this.properties.TemplateList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldListPicker('EmployeeList', {
                  label: 'Select Employee List',
                  selectedList: this.properties.EmployeeList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
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
