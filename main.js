const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("node:path");
const fs = require("fs").promises;
const mammoth = require("mammoth");
const ExcelJS = require("exceljs");
const { table, error } = require("node:console");
const { text } = require("node:stream/consumers");

if (handleSquirrelEvent()) {
  app.quit();
  return;
} else {
  main().catch(console.error);
}

function handleSquirrelEvent() {
  if (process.argv.length === 1) {
    return false;
  }

  const ChildProcess = require("child_process");
  const path = require("path");

  const appFolder = path.resolve(process.execPath, "..");
  const rootAtomFolder = path.resolve(appFolder, "..");
  const updateDotExe = path.resolve(path.join(rootAtomFolder), "Update.exe");
  const exeName = path.basename(process.execPath);

  const spawn = function (command, args) {
    let spawnedProcess, error;
    try {
      spawnedProcess = ChildProcess.spawn(command, args, { detached: true });
    } catch (error) {}
    return spawnedProcess;
  };

  const spawnUpdate = function (args) {
    return spawn(updateDotExe, args);
  };

  const squirrelEvent = process.argv[1];

  switch (squirrelEvent) {
    case "--squirrel-install":
    case "--squirrel-updated":
      spawnUpdate(["--createShortcut", exeName]);
      setTimeout(() => {
        app.quit();
      }, 10000);
      return true;
    case "--squirrel-uninstall":
      spawnUpdate(["--removeShortcut", exeName]);
      setTimeout(() => {
        app.quit();
      }, 10000);
      return true;
    case "--squirrel-firstrun":
    case "--squirrel-obsolete":
      app.quit();
      return true;
  }
}

function parseTableFromHtml(htmlString) {
  const cheerio = require("cheerio");
  const $ = cheerio.load(htmlString);
  const tables = [];
  $("table").each((idx, table) => {
    const data = [];
    $(table)
      .find("tr")
      .each((_, row) => {
        const rowData = [];
        $(row)
          .find("td", "th")
          .each((_, cell) => {
            const colspan = $(cell).attr("colspan")
              ? parseInt($(cell).attr("colspan"))
              : 1;
            const rowspan = $(cell).attr("rowspan")
              ? parseInt($(cell).attr("rowspan"))
              : 1;
            rowData.push({
              text: $(cell).text().trim(),
              colspan,
              rowspan,
            });
          });
        data.push(rowData);
      });
    tables.push(data);
  });
  return tables;
}

const DEFAULT_TASK = {
  level: "220kV",
  type: "典型操作任务",
  sub_type: "xxkV操作任务",
  separation: "未知间隔",
};

const info = {
  name: "xx变",
  tasks: new Map(),
};

const global = {
  workbook: null,
};

function addTask(key, custom) {
  const task = { ...DEFAULT_TASK, ...custom };
  info.tasks.set(key, task);
}

async function getDataFromExample(tableData) {
  info.name = tableData[1][1].trim();
  for (let i = 0; i < tableData.length; ++i) {
    if (tableData[i][0] !== null || tableData[i][1] != null) {
      addTask(tableData[i][4].replace(/\s+/g, ""), {
        ...(tableData[i][0] ? { level: tableData[i][0] } : {}),
        ...(tableData[i][2] ? { type: tableData[i][2] } : {}),
        ...(tableData[i][3] ? { sub_type: tableData[i][3] } : {}),
        ...(tableData[i][5]
          ? {
              separation:
                typeof tableData[i][5] == "string"
                  ? tableData[i][5]
                  : tableData[i][5].result,
            }
          : {}),
      });
    }
  }
}

async function saveTablesToExcel(tablesData) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("sheet 1");
  let ridx = 1;
  const firstRow = worksheet.addRow([
    "变电所最高电压等级",
    "变电所",
    "操作类型",
    "操作子类型",
    "操作任务",
    "适用间隔",
  ]);
  const skip = /^(说明|备注|顺序)/;
  for (let tIndex = 1; tIndex < tablesData.length; tIndex++) {
    let idx = 1;
    for (
      let rowIndex = 0;
      rowIndex < tablesData[tIndex].length;
      rowIndex++, ridx++
    ) {
      const row = tablesData[tIndex][rowIndex];
      const excelRow = worksheet.addRow([]);
      let colIndex = 3;

      let re = /.*操作任务.*(?<!(所用|主|压|接地))变/;
      // console.log(re);
      for (const cell of row) {
        if (cell.text.match(skip)) {
          worksheet.spliceRows(ridx + 1, 1);
          ridx--;
          break;
        } else if (cell.text.length < 5) {
          if (row.length > 1) excelRow.getCell(colIndex + 1).value = cell.text;
          else {
            worksheet.spliceRows(ridx + 1, 1);
            ridx--;
            break;
          }
        } else if (cell.text.match(re)) {
          const taskName = cell.text.replace(re, "").replace(/\s+/g, "");
          let taskInfo = info.tasks.get(taskName) || DEFAULT_TASK;
          excelRow.getCell(1).value = taskInfo.level;
          excelRow.getCell(2).value = info.name;
          excelRow.getCell(3).value = taskInfo.type;
          excelRow.getCell(4).value = taskInfo.sub_type;
          excelRow.getCell(5).value = taskName;
          excelRow.getCell(6).value = taskInfo.separation;
          break;
        } else {
          if (row.length < 2) {
            worksheet.spliceRows(ridx + 1, 1);
            ridx--;
            break;
          }
          excelRow.getCell(colIndex + 1).value = cell.text;
        }

        if (cell.colspan > 1 || cell.rowspan > 1) {
          const startCol = colIndex + 1;
          const endCol = startCol + cell.colspan - 1;
          const startRow = ridx + 1;
          const endRow = startRow + cell.rowspan - 1;
          // worksheet.mergeCells(startRow, startCol, endRow, endCol);
        }
        // colIndex += cell.colspan;
        colIndex += 1;
      }
    }
  }

  worksheet.columns.forEach((column) => {
    column.width = 15;
  });
  global.workbook = workbook;
  return true;
}

function createWindow() {
  const win = new BrowserWindow({
    titleBarStyle: "hidden",
    resizable: false,
    ...(process.platform !== "darwin"
      ? {
          titleBarOverlay: {
            color: "#00000000",
            symbolColor: "#777",
            height: 30,
          },
        }
      : {}),
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  win.loadFile("index.html");

  ipcMain.handle(
    "dialog:saveFile",
    async (e, defaultFileName = "take it.xlsx") => {
      try {
        const { filePath } = await dialog.showSaveDialog(
          BrowserWindow.getFocusedWindow(),
          {
            title: "保存 Excel 文件",
            defaultPath: defaultFileName,
            filters: [{ name: "Excel 文件", extensions: ["xlsx"] }],
          },
        );

        if (filePath) {
          await global.workbook.xlsx.writeFile(filePath);
          console.log(`Excel 文件已成功保存到: ${filePath}`);
          return {
            success: true,
            path: filePath,
          };
        } else {
          console.log("用户取消了保存操作");
          return {
            success: false,
            error: "用户取消保存",
          };
        }
      } catch (e) {
        return {
          success: false,
          error: e.message,
        };
      }
    },
  );

  ipcMain.handle("dialog:openFile", async (e, type) => {
    const result = await dialog.showOpenDialog(win, {
      properties: ["openFile"],
      filters: [{ name: "Word Doc", extensions: type }],
    });
    if (result.canceled) return null;
    return result.filePaths[0];
  });

  ipcMain.handle("word:parse", async (e, filePath) => {
    try {
      const fileBuffer = await fs.readFile(filePath);
      const result = await mammoth.convertToHtml({ path: filePath });
      const tables = parseTableFromHtml(result.value);
      saveTablesToExcel(tables);
      return {
        success: true,
        raw: result.value,
        html: JSON.stringify(tables, null, 2),
      };
    } catch (error) {
      console.log("[Word] Parse failed.", error);
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle("excel:parse", async (e, filePath) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);

      const worksheet = workbook.getWorksheet(1);
      if (!worksheet) {
        return { success: false, error: "Excel 文件中没有找到工作表" };
      }

      const rows = [];
      worksheet.eachRow((row, rowNumber) => {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData.push(cell.value);
        });
        rows.push(rowData);
      });

      const merges = worksheet.model.merges || [];
      console.log("合并单元格区域:", merges);

      getDataFromExample(rows);
      return { success: true, data: rows, merges: merges };
    } catch (error) {
      console.error("读取 Excel 失败:", error);
      return { success: false, error: error.message };
    }
  });
}

async function main() {
  await app.whenReady();

  // ipcMain.handle("ping", () => {
  //   return "pong";
  // });

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });

  app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
      app.quit();
    }
  });

  createWindow();
}
