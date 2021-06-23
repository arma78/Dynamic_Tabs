declare interface IDynamicTabsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TitleFieldLabel: string;
  ListName:string;
  ListNameFieldLabel: string;
  termsetNameOrIDFieldLabel: string;
  fieldNameFiedLabel: string;
}

declare module 'DynamicTabsWebPartStrings' {
  const strings: IDynamicTabsWebPartStrings;
  export = strings;
}
