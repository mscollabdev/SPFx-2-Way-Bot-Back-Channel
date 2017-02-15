declare interface IEchoBotStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'echoBotStrings' {
  const strings: IEchoBotStrings;
  export = strings;
}
