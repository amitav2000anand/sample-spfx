declare interface IBubbleChatApplicationCustomizerStrings {
  Title: string;
  DefaultButtonLabel: string;
  DefaultBotName: string;
}

declare module "BubbleChatApplicationCustomizerStrings" {
  const strings: IBubbleChatApplicationCustomizerStrings;
  export = strings;
}
