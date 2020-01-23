declare interface IDirectoryWebPartStrings {
  DropDownPlaceLabelMessage: string;
  DropDownPlaceHolderMessage: string;
  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  SearchFirstNameLabel: string;
  ShowSortLabel: string;
  DirectoryMessage: string;
  DescriptionFieldLabel: string;
  FirstName: string;
  LastName: string;
  Department: string;
  Location: string;
  JobTitle: string;
  defaultSortLabel: string;
  lblSearch: string;
}

declare module 'DirectoryWebPartStrings' {
  const strings: IDirectoryWebPartStrings;
  export = strings;
}
