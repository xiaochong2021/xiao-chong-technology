import { dialog, ipcMain, IpcMainInvokeEvent, IpcMainEvent, BrowserWindow, Notification } from 'electron';
import XLSX from 'xlsx';
import { RegParseModel } from './reg-parse-model';
import path from 'path';

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
  const workBook = XLSX.readFile(filePath);
  const sheetName = workBook.SheetNames[0];
  const sheet = workBook.Sheets[sheetName];
  const firstRow = XLSX.utils.sheet_to_json(sheet)[0];
  return Object.keys(firstRow as object);
}

function matchTextItem(searchText: string, config: RegParseModel, regSheet: XLSX.WorkSheet, regNumList: RegExpMatchArray, regStrList: Array<string>) : string {
  const regMatchSet:Set<string> = new Set(); //匹配到正则组名称
  let searchTextSplitList:string[];
  if (config.splitText) {
    searchTextSplitList = String(searchText).split(config.splitText);
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
          regLogicMap.set(regNum, new RegExp(regCell.v, config.isCaseSensitive ? 'i' : '').test(searchTextItem));
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
  contentSheet[regAddr] = {v:regText};
}

ipcMain.handle('dialog:openFile', handleFileOpen);

ipcMain.handle('parseColumns', handleParseColumn);



export function postRenderMessage(messageWindow: BrowserWindow) {
  ipcMain.on('executeRegMatch', handleParse);
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
        continue;
      }
      const contentAddr = XLSX.utils.encode_col(config.contentColumnSelect) + XLSX.utils.encode_row(contentRowStart);

      let content = contentSheet[contentAddr]?.v;
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
      const resultStr = matchTextItem(content, config, regSheet, regNumList as RegExpMatchArray, regStrList);
      contentSheetAdd(contentSheet, contentRowStart, contentColEndNext, resultStr);
      const progressRate = contentRowStart/contentRowEnd;
      messageWindow.setProgressBar(progressRate);
      messageWindow?.webContents.send('updateState', {state: 'doing', progress: (progressRate*100).toPrecision(2)});
    }
    messageWindow.setProgressBar(0);
    messageWindow?.webContents.send('updateState', {state: 'done', progress: '100'});
    const resultWb = XLSX.utils.book_new();
    const range = contentSheet['!ref']?.split(':') || [];
    if (range.length > 0) {
      contentSheet['!ref'] = range[0] + ':' + XLSX.utils.encode_col(contentColEndNext) + XLSX.utils.encode_row(--contentRowStart);
    }
    XLSX.utils.book_append_sheet(resultWb, contentSheet, '正则匹配结果');
    const contentFileName = path.basename(config.contentFilePath).split('.')[0];
    const outFilePath = path.resolve(config.contentFilePath, '..', contentFileName + '_正则匹配结果_' + Date.now() + '.xlsx')
    XLSX.writeFileXLSX(resultWb, outFilePath);

    const notification = new Notification({
      title: '晓蟲技术支持',
      body: '批量正则匹配已完成，请到文本文件所在目录，查看匹配结果文件！',
      silent: true,
    });
    // notification.on('click', () => {
    //   const hasFile = fs.existsSync(outFilePath);
    //   if (hasFile) {
    //     fs.open(outFilePath, 'r+', () => {
    //       new Notification({
    //         title: '晓蟲技术支持',
    //         body: '打开文件失败！',
    //       }).show();
    //     });
    //   }
    // });
    notification.show();
  }
}





