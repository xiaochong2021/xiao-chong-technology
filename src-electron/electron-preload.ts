import { contextBridge, IpcRendererEvent, ipcRenderer } from 'electron';
import { RegParseModel } from './modules/reg-parse-model';

type cb = (event: IpcRendererEvent, res:{state: string, progress: string}) => void

declare global {
  interface Window  {
    electronAPI: {
      openFile: () => Promise<string>,
      parseColumns: (filePath: string) => Promise<[]>,
      executeRegMatch: (config: RegParseModel) => void,
      onUpdateState: (callback: cb) => void
    }
  }
}

contextBridge.exposeInMainWorld('electronAPI',{
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  parseColumns: (filePath: string) => ipcRenderer.invoke('parseColumns', filePath),
  executeRegMatch: (config: RegParseModel) => ipcRenderer.send('executeRegMatch', config),
  onUpdateState: (callback: cb) => ipcRenderer.on('updateState', callback)
});

/**
 * 根据配置过滤原文本
 */
function contentFilter(content: string, config: RegParseModel) {
  content = String(content);
  if (config.isFilterUserName) { // 剔除@用户名
    content = content.replace(/@[\u4e00-\u9fa5a-zA-Z0-9_-]{4,30}/g, '');
  }
  if (config.isFilterTopic) { // 剔除#话题#
    content = content.replace(/#([^#]+)#/g, '');
  }
  if (config.isFilterSpecTopic) { // 剔除#特殊格式话题
    content = content.replace(/#\S+/g, '');
  }
  if (config.isFilterEmoticon) { // 剔除表情
    content = content.replace(/\[[^\[\]]+\]/g, '');
  }
  if (config.isFilterURL) { // 剔除链接
    content = content.replace(/(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?/g, '');
  }
  return content;
}

function matchRegRow(content: string, regSheetData: Array<Array<string>>, config: RegParseModel, regNumList: RegExpMatchArray | Array<string>, regStrList: Array<string>) {
  const regMatchSet:Set<string> = new Set(); //匹配到正则组名称
  let contentSplitList:string[];

  if (config.splitText) {
    contentSplitList = String(content).split(new RegExp(config.splitText));
  } else {
    contentSplitList = [String(content)]
  }

  const rowData = [];
  // const headerRow = regSheetData[0];
  let regFlags = 'g';
  regFlags += config.isNotCaseSensitive ? 'i' : '';

  regSheetData.forEach((regRow, regIndex) => {
    if (regIndex !== 0) { // 第一行为表头不需要处理
      const regLogicMap:Map<string,  Array<RegExpMatchArray | null> | null> = new Map(); //逻辑码表转换辅助映射
      let showCell = false;

      contentSplitList.forEach(searchTextItem => {
        regNumList.forEach(regNum => {
          const matchResult = searchTextItem.match(new RegExp('.{0,15}(' + regRow[Number.parseInt(regNum)] + ').{0,15}', regFlags));
          if (!regLogicMap.has(regNum)) {
            regLogicMap.set(regNum, [matchResult]);
          } else {
            regLogicMap.get(regNum)?.push(matchResult)
          }
        });

        const transformRegLogicCode = regStrList.map(regStr => {
          if (regLogicMap.has(regStr)) {
            const regLogicItem = regLogicMap.get(regStr);
            return regLogicItem && regLogicItem.some(entryItem => entryItem !== null);
          } else {
            return regStr;
          }
        }).join(' ');

        try {
          /* eslint-disable-next-line */
          if (eval(transformRegLogicCode)) {
            const regTheme = regRow[config.regColumnSelect];
            regMatchSet.add(regTheme);
            showCell = true;
          }
        } catch (e) {
          regMatchSet.add('转换逻辑码表出错');
        }
      });

      const matchTextList = [];
      for (const regLogicMapEntry of regLogicMap.entries()) {
        if (regLogicMapEntry[1] && regLogicMapEntry[1].some(entryItem => entryItem !== null) && showCell) {
          matchTextList.push(/*headerRow[Number.parseInt(regLogicMapEntry[0])] + '：' +*/ regLogicMapEntry[1].map(regMatchItem => {
            if (regMatchItem) {
              return regMatchItem.join(' ')
            } else {
              return '';
            }
          }).filter(str => !!str).join('  >>>  ').trim())
        }
      }

      rowData.push({
        v: matchTextList.join('\n').trim(),
        s: {
          alignment: {
            wrapText: true,
            vertical: 'top',
          }
        }
      })
    }
  });

  rowData.unshift({
    v: Array.from(regMatchSet).join('，'),
    s: {
      alignment: {
        wrapText: true,
        vertical: 'top',
      }
    }
  });
  return rowData;
}

function matchTextItem(searchText: string, config: RegParseModel, regSheetData: Array<Array<string>>, regNumList: RegExpMatchArray| Array<string>, regStrList: Array<string>) : string {
  const regMatchSet:Set<string> = new Set(); //匹配到正则组名称
  let searchTextSplitList:string[];
  if (config.splitText) {
    searchTextSplitList = String(searchText).split(new RegExp(config.splitText));
  } else {
    searchTextSplitList = [String(searchText)]
  }

  regSheetData.forEach((regRow, regIndex) => {
    if (regIndex !== 0) {
      searchTextSplitList.forEach(searchTextItem => {
        const regLogicMap:Map<string, boolean> = new Map(); //逻辑码表转换辅助映射
        regNumList.forEach(regNum => {
          regLogicMap.set(regNum, new RegExp(regRow[Number.parseInt(regNum)], config.isNotCaseSensitive ? 'i' : '').test(searchTextItem));
        });

        const transformRegLogicCode = regStrList.map(regStr => {
          if (regLogicMap.has(regStr)) {
            return String(regLogicMap.get(regStr));
          } else {
            return regStr;
          }
        }).join(' ');
        try {
          /* eslint-disable-next-line */
          if (eval(transformRegLogicCode)) {
            const regTheme = regRow[config.regColumnSelect];
            regMatchSet.add(regTheme);
          }
        } catch (e) {
          regMatchSet.add('转换逻辑码表出错');
        }
      });
    }
  });
  return Array.from(regMatchSet).join('，');
}

ipcRenderer.on('work', (event, data: string[][], regSheetData: Array<Array<string>>, config: RegParseModel, regNumList: RegExpMatchArray | Array<string>, regStrList: Array<string>) => {
  data.forEach((row) => {
    const contentValue = row[config.contentColumnSelect];
    const filterContent = contentFilter(contentValue, config);
    const result: any[] = [
      {
        v: contentValue,
        s: {
          alignment: {
            vertical: 'top',
          }
        }
      }
    ];
    if (config.isSearchText) {
      result.push(...matchRegRow(filterContent, regSheetData, config, regNumList || [], regStrList));
    } else {
      result.push({
        v: matchTextItem(filterContent, config, regSheetData, regNumList || [], regStrList),
        s: {
          alignment: {
            wrapText: true,
            vertical: 'top',
          }
        }
      })
    }
    ipcRenderer.send('workerResult', result)
  });
  window.close();
})




