const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const os = require('os');

const homeDir = os.homedir();
const desktopDir = `${homeDir}/Desktop`;

function createWindow() {
  // Create the browser window.
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
    },
  });

  // and load the index.html of the app.
  win.loadFile('html/index.html');

  // Open the DevTools.
  win.webContents.openDevTools();

  // let menu = Menu.buildFromTemplate([
  //   {
  //     label: 'Pages',
  //     submenu: [
  //       {
  //         label: 'Business report',
  //         click() {
  //           shell.openExternal('https://github.com/vfxpipeline/system-info/blob/master/main.js');
  //         },
  //       },
  //     ],
  //   },
  // ]);

  // Menu.setApplicationMenu(menu);
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(createWindow);

// Quit when all windows are closed.
app.on('window-all-closed', () => {
  // On macOS it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  // On macOS it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

ipcMain.on('open-business-report', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'csv',
          extensions: ['csv'],
        }
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-businessReport', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-kong', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'xlsx',
          extensions: ['xlsx'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-kong', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-kong-inventory', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'csv',
          extensions: ['csv'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-kong-inventory', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-formulaSports', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'xlsx',
          extensions: ['xlsx'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-formulaSports', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-formulaSports-inventory', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'csv',
          extensions: ['csv'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-formulaSports-inventory', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-rapidLoss', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'xlsx',
          extensions: ['xlsx'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-rapidLoss', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-rapidLoss-inventory', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'csv',
          extensions: ['csv'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-rapidLoss-inventory', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-skinPhysics', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'xlsx',
          extensions: ['xlsx'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-skinPhysics', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-skinPhysics-inventory', (event) => {
  dialog
    .showOpenDialog({
      properties: ['openFile'],
      filters: [
        {
          name: 'csv',
          extensions: ['csv'],
        },
      ],
      defaultPath: desktopDir,
    })
    .then((result) => {
      if (result.canceled === false) {
        event.sender.send('selected-skinPhysics-inventory', result);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});
