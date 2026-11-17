const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('desktop', {
  async invoke(action, argument) {
    const response = await ipcRenderer.invoke('sync:invoke', action, argument);
    if (response.error) throw new Error(response.error);
    return response.value;
  },
  filePath: file => webUtils.getPathForFile(file),
});
