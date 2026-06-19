declare interface IRemoveFilesSpoCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'RemoveFilesSpoCommandSetStrings' {
  const strings: IRemoveFilesSpoCommandSetStrings;
  export = strings;
}
