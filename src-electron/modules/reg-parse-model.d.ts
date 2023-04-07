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
  isNotCaseSensitive: boolean,
  isFilterEmoticon: boolean,
  isFilterURL: boolean,
  isSearchText: boolean
}
