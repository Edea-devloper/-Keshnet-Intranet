export interface IKeshnetPhonebookProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  PhoneBookList: string
  userDisplayName: string;
  selectedColumns: string[];
  selectedColumnsCard: string[];
  maxRows?: number;
  context: any;
  tabConfigs:any;
  listColumns?: { key: string; text: string }[]; 
  ImageLibrary:string;
  AssetFolderPath:string;
  selectordercolumn:string;
}
