import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IKeshnetGalleryProps {
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  _GalleryLists: string;
  context:WebPartContext;
  ImageLibrary:string;
  TemplateList:string;
  EmployeeList:string;
  AssetFolderPath:string;
}
