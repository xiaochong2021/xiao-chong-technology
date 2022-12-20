export interface RegParseModel {
  regFilePath: string,
  contentFilePath: string,
  regColumnSelect: number,
  contentColumnSelect: number,
  logicCode: string
  splitText?: string,
  isFilterUserName: boolean,
  isFilterTopic: boolean,
  isFilterSpecTopic: boolean,
  isCaseSensitive: boolean,
  isFilterEmoticon: boolean,
  isFilterURL: boolean,
}
