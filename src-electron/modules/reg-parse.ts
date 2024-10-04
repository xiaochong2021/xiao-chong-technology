import { dialog, ipcMain, IpcMainInvokeEvent, IpcMainEvent, BrowserWindow, Notification } from 'electron';
import XLSX from 'xlsx-js-style';
import { RegParseModel } from './reg-parse-model';
import path from 'path';
import { getFirstRow } from 'app/src-electron/utils/xlsxUtil';

const PROCESS_NUM = 10;

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


ipcMain.handle('dialog:openFile', handleFileOpen);

ipcMain.handle('parseColumns', handleParseColumn);

const createWindowWorker = (data: string[][], regSheetData: Array<Array<string>>, config: RegParseModel, regNumList: RegExpMatchArray | Array<string>, regStrList: Array<string>) => {
  const workerWin = new BrowserWindow({
    show: false,
    webPreferences: {
      contextIsolation: true,
      preload: path.resolve(__dirname, process.env.QUASAR_ELECTRON_PRELOAD),
    },
  });

  workerWin.loadURL(process.env.APP_URL);

  workerWin.webContents?.send('work', data, regSheetData, config, regNumList, regStrList);
}

const setProgressFinished = (messageWindow: BrowserWindow) => {
  messageWindow?.webContents.send('updateState', {state: 'done', progress: '100'});
  messageWindow.setProgressBar(0);
  const notification = new Notification({
    title: '文本智能识别工具',
    body: '批量正则匹配已完成，请到文本文件所在目录，查看匹配结果文件！',
    silent: true,
  });
  notification.show();
}



export function postRenderMessage(messageWindow: BrowserWindow) {
  ipcMain.on('executeRegMatch', handleExecuteRegMatch);

  function handleExecuteRegMatch(event: IpcMainEvent, config: RegParseModel) {
    const regWB = XLSX.readFile(path.resolve(config.regFilePath));
    const contentWB = XLSX.readFile(path.resolve(config.contentFilePath));
    const regSheetData = <Array<Array<string>>> XLSX.utils.sheet_to_json(regWB.Sheets[regWB.SheetNames[0]], {header: 1});
    const contentSheetData = <Array<Array<string>>> XLSX.utils.sheet_to_json(contentWB.Sheets[contentWB.SheetNames[0]], {header: 1});

    const regNumList = config.logicCode.match(/\d+/g); //对逻辑码表的数字进行切分
    const regStrList = config.logicCode.split(/\b(?=\d+)|(?<=\d+)\b/); //对逻辑码表的所有字符进行切分
    const exportData: Array<Array<object>> = [];
    const totalRow = contentSheetData.length;

    if (totalRow) {
      const firstRow = contentSheetData[0];
      const contentValue = firstRow[config.contentColumnSelect];
      const exportHeaderRow = [
        {
          v: contentValue,
          t: 's',
          s: headerStyle
        }
      ];
      if (config.isSearchText) {
        exportHeaderRow.push(...regSheetData.map((item) => ({
          v: item[config.regColumnSelect],
          t: 's',
          s: headerStyle
        })));
      } else {
        exportHeaderRow.push({
          v: regSheetData[0][config.regColumnSelect],
          t: 's',
          s: headerStyle,
        });
      }
      exportData.push(exportHeaderRow);

      const PROCESS_SIZE = Math.ceil(totalRow / PROCESS_NUM);
      for (let processIndex = 0; processIndex < PROCESS_NUM; processIndex++) {
        if (processIndex === 0) {
          const firstBatch = contentSheetData.slice(1, PROCESS_SIZE);
          createWindowWorker(firstBatch, regSheetData, config, regNumList || [], regStrList);
        } else {
          const batch = contentSheetData.slice(processIndex * PROCESS_SIZE, (processIndex + 1) * PROCESS_SIZE);
          createWindowWorker(batch, regSheetData, config, regNumList || [], regStrList);
        }
      }
    }

    let finishedNum = 1;
    ipcMain.on('workerResult', (event, result) => {
      exportData.push((result));
      finishedNum++;
      const progressRate = finishedNum/totalRow;
      messageWindow.setProgressBar(progressRate);
      const progressRate2 = (progressRate*100).toPrecision(2);
      messageWindow?.webContents.send('updateState', {state: 'doing', progress: progressRate2});

      if (finishedNum === totalRow) {
        const resultWb = XLSX.utils.book_new();
        const resultSheet = XLSX.utils.aoa_to_sheet(exportData);
        resultSheet['!rows'] = [{hpx: 30}];
        resultSheet['!cols'] = exportData[0].map(() => ({wpx: 200}))

        XLSX.utils.book_append_sheet(resultWb, resultSheet, '正则匹配结果');
        const contentFileName = path.basename(config.contentFilePath).split('.')[0];
        const outFilePath = path.resolve(config.contentFilePath, '..', contentFileName + '_正则匹配结果_' + Date.now() + '.xlsx')
        XLSX.writeFile(resultWb, outFilePath);

        setProgressFinished(messageWindow);
      }
    });
  }
}





