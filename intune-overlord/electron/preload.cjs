const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('intuneOverlordDesktop', {
  isDesktop: true,
  getRuntimeConfig: () => ipcRenderer.invoke('runtime-config:get'),
  resetConfig: () => ipcRenderer.invoke('config:reset'),
  runTenantSetup: (script) => ipcRenderer.invoke('tenant-setup:run', script),
})
