export interface IKeshnetPhonebookProps {
  description: string;
  isDarkTheme: boolean;
  themeFont?: string;
  environmentMessage: string;
  hasTeamsContext: boolean;
  PhoneBookList: string
  userDisplayName: string;
  selectedColumnsFirst: string[];
  selectedColumnsSecond: string[];
  selectedColumnsThird: string[];
  selectedColumnsCardFirstTab: string[];
  selectedColumnsCardSecondTab: string[];
  selectedColumnsCardThirdTab: string[];
  maxRows?: number;
  context: any;
  tabConfigs:any;
  listColumns?: { key: string; text: string }[]; 
  ImageLibrary:string;
  AssetFolderPath:string;
  selectordercolumn:string;
}
