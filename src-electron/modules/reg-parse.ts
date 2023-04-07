import { dialog, ipcMain, IpcMainInvokeEvent, IpcMainEvent, BrowserWindow, Notification } from 'electron';
import XLSX from 'xlsx-js-style';
import { RegParseModel } from './reg-parse-model';
import path from 'path';
import { getFirstRow } from 'app/src-electron/utils/xlsxUtil';


const headerStyle = {
  font: { bold: true, color: {rgb: 'FFFFFF'}},
  fill: { fgColor: { rgb: '494529' }},
  alignment: { vertical: 'center', horizontal: 'center'},
};

/**
 * @description 打开文件选择框，返回选中文件路径
 */
async function handleFileOpen() {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    filters: [
      { name: 'Excel', extensions: ['xls', 'xlsx', 'csv'] },
    ]
  });
  if (canceled) {
    return
  } else {
    return filePaths[0]
  }
}

/**
 * @description 解析excel，获取其列数组
 */
async function handleParseColumn(event: IpcMainInvokeEvent, filePath: string) {
  return getFirstRow(filePath);
}

function matchTextItem(searchText: string, config: RegParseModel, regSheet: XLSX.WorkSheet, regNumList: RegExpMatchArray, regStrList: Array<string>) : string {
  const regMatchSet:Set<string> = new Set(); //匹配到正则组名称
  let searchTextSplitList:string[];
  if (config.splitText) {
    searchTextSplitList = String(searchText).split(new RegExp(config.splitText));
  } else {
    searchTextSplitList = [String(searchText)]
  }
  const regRange = XLSX.utils.decode_range(<string>regSheet['!ref']);
  let regRowStart = regRange.s.r;
  const  regRowEnd = regRange.e.r;
  for(; regRowStart <= regRowEnd; regRowStart++) {
    if (regRowStart === 0) continue; //第一行为表头，不处理
    searchTextSplitList.forEach(searchTextItem => {
      const regLogicMap:Map<string, boolean> = new Map(); //逻辑码表转换辅助映射
      regNumList.forEach(regNum => {
        if (!regLogicMap.has(regNum)) {
          const regAddr = XLSX.utils.encode_col(Number.parseInt(regNum)) + XLSX.utils.encode_row(regRowStart);
          const regCell = regSheet[regAddr];
          regLogicMap.set(regNum, new RegExp(regCell.v, config.isNotCaseSensitive ? 'i' : '').test(searchTextItem));
        }
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
          const regTheme = regSheet[XLSX.utils.encode_col(config.regColumnSelect) + XLSX.utils.encode_row(regRowStart)].v;
          regMatchSet.add(regTheme);
        }
      } catch (e) {
        regMatchSet.add('转换逻辑码表出错');
      }
    });
  }
  return Array.from(regMatchSet).join('，');
}

/**
 * @description 在原文本表格最后一列添加正则命中问题描述
 */
function contentSheetAdd(contentSheet:XLSX.WorkSheet, contentIndex: number , regIndex: number, regText: string) {
  const regAddr = XLSX.utils.encode_col(regIndex) + XLSX.utils.encode_row(contentIndex);
  contentSheet[regAddr] = {
    v: regText
  };
}

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

/**
 * 单行正则匹配
 */
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



ipcMain.handle('dialog:openFile', handleFileOpen);

ipcMain.handle('parseColumns', handleParseColumn);



export function postRenderMessage(messageWindow: BrowserWindow) {
  ipcMain.on('executeRegMatch', handleRegMatch);

  function handleRegMatch(event: IpcMainEvent, config: RegParseModel) {
    if (config.isSearchText) {
      handleExecuteRegMatch(event, config);
    } else {
      handleParse(event, config);
    }
    const notification = new Notification({
      title: '晓蟲技术支持',
      body: '批量正则匹配已完成，请到文本文件所在目录，查看匹配结果文件！',
      silent: true,
    });
    notification.show();

    messageWindow?.webContents.send('updateState', {state: 'doing', progress: '100'});
    messageWindow?.webContents.send('updateState', {state: 'done', progress: '100'});
    messageWindow.setProgressBar(0);
  }

  function handleExecuteRegMatch(event: IpcMainEvent, config: RegParseModel) {
    const regWB = XLSX.readFile(path.resolve(config.regFilePath));
    const contentWB = XLSX.readFile(path.resolve(config.contentFilePath));
    const regSheetData = <Array<Array<string>>> XLSX.utils.sheet_to_json(regWB.Sheets[regWB.SheetNames[0]], {header: 1});
    const contentSheetData = <Array<Array<string>>> XLSX.utils.sheet_to_json(contentWB.Sheets[contentWB.SheetNames[0]], {header: 1});

    const regNumList = config.logicCode.match(/\d+/g); //对逻辑码表的数字进行切分
    const regStrList = config.logicCode.split(/\b(?=\d+)|(?<=\d+)\b/); //对逻辑码表的所有字符进行切分
    const exportData: Array<Array<object>> = [];
    const totalRow = contentSheetData.length;

    contentSheetData.forEach((row, rowIndex) => {
      const contentValue = row[config.contentColumnSelect];
      if (rowIndex === 0) { // 添加表头
        exportData.push([
          {
            v: contentValue,
            t: 's',
            s: headerStyle
          },
          ...regSheetData.map((item) => ({
            v: item[config.regColumnSelect],
            t: 's',
            s: headerStyle
          }))
        ])
      } else {
        const filterContent = contentFilter(contentValue, config);
        exportData.push([
          {
            v: contentValue,
            s: {
              alignment: {
                vertical: 'top',
              }
            }
          },
          ...matchRegRow(filterContent, regSheetData, config, regNumList || [], regStrList)
        ]);
      }
      const progressRate = rowIndex/totalRow;
      messageWindow.setProgressBar(progressRate);
      const progressRate2 = (progressRate*100) > 99 ? '99' : (progressRate*100).toPrecision(2);

      messageWindow?.webContents.send('updateState', {state: 'doing', progress: progressRate2});
    });

    const resultWb = XLSX.utils.book_new();
    const resultSheet = XLSX.utils.aoa_to_sheet(exportData);
    resultSheet['!rows'] = [{hpx: 30}];
    resultSheet['!cols'] = exportData[0].map(() => ({wpx: 200}))

    XLSX.utils.book_append_sheet(resultWb, resultSheet, '正则匹配结果');
    const contentFileName = path.basename(config.contentFilePath).split('.')[0];
    const outFilePath = path.resolve(config.contentFilePath, '..', contentFileName + '_正则匹配结果_' + Date.now() + '.xlsx')
    XLSX.writeFile(resultWb, outFilePath);
  }

  function handleParse(event: IpcMainEvent, config: RegParseModel) {
    const regWorkBook = XLSX.readFile(path.resolve(config.regFilePath) );
    const contentWorkBook = XLSX.readFile(path.resolve(config.contentFilePath));
    const regSheet = regWorkBook.Sheets[regWorkBook.SheetNames[0]];
    const contentSheet = contentWorkBook.Sheets[contentWorkBook.SheetNames[0]];
    const contentRange = XLSX.utils.decode_range(<string>contentSheet['!ref']);
    let  contentRowStart = contentRange.s.r;
    const contentRowEnd = contentRange.e.r;
    const contentColEndNext = contentRange.e.c + 1;

    const regNumList = config.logicCode.match(/\d+/g); //对逻辑码表的数字进行切分
    const regStrList = config.logicCode.split(/\b(?=\d+)|(?<=\d+)\b/); //对逻辑码表的所有字符进行切分

    for(;contentRowStart <= contentRowEnd; contentRowStart++) {

      if (contentRowStart === 0) {
        // 第一行为表头，在最后一列添加正则命中问题
        contentSheetAdd(contentSheet, contentRowStart, contentColEndNext, regSheet[XLSX.utils.encode_col(config.regColumnSelect) + XLSX.utils.encode_row(0)].v);
        contentSheet['!rows'] = [{hpx: 30}];
        contentSheet['!cols'] = Array(contentColEndNext + 1).toString().split(',').map(() => ({wpx: 200}))
        for (let i = 0; i <= contentColEndNext; i++) {
          contentSheet[XLSX.utils.encode_cell({c:i, r: contentRowStart})].s = headerStyle;
        }
        continue;
      }
      const contentAddr = XLSX.utils.encode_col(config.contentColumnSelect) + XLSX.utils.encode_row(contentRowStart);
      const content = contentFilter(contentSheet[contentAddr]?.v, config);

      const resultStr = matchTextItem(content, config, regSheet, regNumList as RegExpMatchArray, regStrList);
      contentSheetAdd(contentSheet, contentRowStart, contentColEndNext, resultStr);
      const progressRate = contentRowStart/contentRowEnd;
      messageWindow.setProgressBar(progressRate);
      messageWindow?.webContents.send('updateState', {state: 'doing', progress: (progressRate*100).toPrecision(2)});
    }

    const resultWb = XLSX.utils.book_new();
    const range = contentSheet['!ref']?.split(':') || [];
    if (range.length > 0) {
      contentSheet['!ref'] = range[0] + ':' + XLSX.utils.encode_col(contentColEndNext) + XLSX.utils.encode_row(--contentRowStart);
    }
    XLSX.utils.book_append_sheet(resultWb, contentSheet, '正则匹配结果');
    const contentFileName = path.basename(config.contentFilePath).split('.')[0];
    const outFilePath = path.resolve(config.contentFilePath, '..', contentFileName + '_正则匹配结果_' + Date.now() + '.xlsx')
    XLSX.writeFile(resultWb, outFilePath);
  }
}





