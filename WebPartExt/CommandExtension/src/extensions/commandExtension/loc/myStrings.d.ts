declare interface ICommandExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CommandExtensionCommandSetStrings' {
  const strings: ICommandExtensionCommandSetStrings;
  export = strings;
}
