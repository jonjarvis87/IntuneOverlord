const path = require('node:path')
const http = require('node:http')
const fs = require('node:fs')
const { exec } = require('node:child_process')
const express = require('express')
const { app, BrowserWindow, ipcMain, shell } = require('electron')
const { PublicClientApplication } = require('@azure/msal-node')

const isDev = Boolean(process.env.VITE_DEV_SERVER_URL)
const desktopPort = Number(process.env.INTUNE_OVERLORD_PORT || 4783)

let mainWindow = null
let localServer = null

const graphResourceAppId = '00000003-0000-0000-c000-000000000000'
const bootstrapClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
const bootstrapScopes = [
  'Application.ReadWrite.All',
  'AppRoleAssignment.ReadWrite.All',
  'DelegatedPermissionGrant.ReadWrite.All',
  'Directory.Read.All',
]
const requiredGraphDelegatedScopes = [
  'DeviceManagementConfiguration.ReadWrite.All',
  'DeviceManagementManagedDevices.Read.All',
  'DeviceManagementScripts.ReadWrite.All',
  'Group.Read.All',
]

function readRuntimeConfig() {
  const appRoot = path.join(__dirname, '..')
  const envPath = path.join(appRoot, '.env')
  const envValues = {
    clientId: '',
    tenantId: 'organizations',
  }

  if (!fs.existsSync(envPath)) {
    return envValues
  }

  const content = fs.readFileSync(envPath, 'utf8')
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith('#')) {
      continue
    }

    const separatorIndex = trimmed.indexOf('=')
    if (separatorIndex === -1) {
      continue
    }

    const key = trimmed.slice(0, separatorIndex).trim()
    const value = trimmed.slice(separatorIndex + 1).trim()

    if (key === 'VITE_AZURE_CLIENT_ID') {
      envValues.clientId = value
    }

    if (key === 'VITE_AZURE_TENANT_ID' && value) {
      envValues.tenantId = value
    }
  }

  return envValues
}

function updateEnvFile(envValues) {
  const appRoot = path.join(__dirname, '..')
  const envPath = path.join(appRoot, '.env')
  const existingContent = fs.existsSync(envPath) ? fs.readFileSync(envPath, 'utf8') : ''
  const existingLines = existingContent ? existingContent.split(/\r?\n/) : []
  const filteredLines = existingLines.filter(
    (line) => !/^VITE_AZURE_CLIENT_ID=/.test(line) && !/^VITE_AZURE_TENANT_ID=/.test(line),
  )

  if (envValues.VITE_AZURE_CLIENT_ID) {
    filteredLines.push(`VITE_AZURE_CLIENT_ID=${envValues.VITE_AZURE_CLIENT_ID}`)
  }

  if (envValues.VITE_AZURE_TENANT_ID) {
    filteredLines.push(`VITE_AZURE_TENANT_ID=${envValues.VITE_AZURE_TENANT_ID}`)
  }

  fs.writeFileSync(envPath, `${filteredLines.filter(Boolean).join('\n')}\n`, 'utf8')
}

function escapeODataValue(value) {
  return value.replaceAll("'", "''")
}

async function graphRequest(accessToken, method, endpoint, body) {
  const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: body ? JSON.stringify(body) : undefined,
  })

  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Graph ${method} ${endpoint} failed: ${response.status} ${errorText}`)
  }

  if (response.status === 204) {
    return undefined
  }

  return response.json()
}

function openExternalBrowser(url) {
  return new Promise((resolve) => {
    if (process.platform === 'win32') {
      // URL is inside double quotes so & is treated literally — no escaping needed
      exec(`cmd /c start "" "${url}"`, () => resolve())
    } else {
      shell.openExternal(url).then(resolve).catch(resolve)
    }
  })
}

async function runDirectOnboarding({ tenantId, appName, redirectUri }) {
  const logs = []
  const effectiveTenant = tenantId?.trim()
  const effectiveName = appName?.trim() || 'Intune Overlord'
  const effectiveRedirect = redirectUri?.trim() || `http://127.0.0.1:${desktopPort}`

  if (!effectiveTenant) {
    return {
      success: false,
      exitCode: -1,
      stdout: '',
      stderr: '',
      error: 'Tenant ID is required.',
      envUpdated: false,
      envValues: {},
    }
  }

  let mainWindowWasVisible = false
  if (mainWindow && !mainWindow.isDestroyed() && mainWindow.isVisible() && !mainWindow.isMinimized()) {
    mainWindowWasVisible = true
    mainWindow.minimize()
  }

  try {
    logs.push(`[1/6] Authenticating with tenant ${effectiveTenant}...`)
    logs.push('       A browser window will open — sign in as a Global Administrator.')

    const pca = new PublicClientApplication({
      auth: {
        clientId: bootstrapClientId,
        authority: `https://login.microsoftonline.com/${effectiveTenant}`,
        redirectUri: 'http://localhost',
      },
    })

    let tokenResult
    try {
      tokenResult = await pca.acquireTokenInteractive({
        scopes: bootstrapScopes,
        prompt: 'select_account',
        openBrowser: async (url) => {
          await openExternalBrowser(url)
        },
        successTemplate:
          '<html><body style="font-family:sans-serif;padding:2rem;background:#0a0b10;color:#ffe81f"><h2>Authentication complete</h2><p style="color:#d8ca8a">You can close this tab and return to Intune Overlord.</p></body></html>',
        errorTemplate:
          '<html><body style="font-family:sans-serif;padding:2rem;background:#0a0b10;color:#ffe81f"><h2>Authentication failed</h2><p style="color:#ffd2da">An error occurred. Close this tab and try again.</p></body></html>',
      })
    } catch (authError) {
      throw new Error(
        `Authentication failed: ${authError instanceof Error ? authError.message : String(authError)}`,
      )
    }

    const accessToken = tokenResult.accessToken
    logs.push('[1/6] Authenticated successfully.')

    // Step 2: Find or create app registration
    logs.push(`[2/6] Looking up app registration "${effectiveName}"...`)
    const escapedName = escapeODataValue(effectiveName)
    const appSearch = await graphRequest(
      accessToken,
      'GET',
      `/applications?$filter=displayName eq '${escapedName}'&$select=id,appId,displayName,requiredResourceAccess`,
    )

    let appRecord = (appSearch.value ?? [])[0]
    let appCreated = false
    if (!appRecord) {
      logs.push('[2/6] No existing registration found — creating...')
      appRecord = await graphRequest(accessToken, 'POST', '/applications', {
        displayName: effectiveName,
        signInAudience: 'AzureADMyOrg',
        isFallbackPublicClient: true,
        spa: {
          redirectUris: [effectiveRedirect],
        },
        publicClient: {
          redirectUris: [effectiveRedirect],
        },
      })
      appCreated = true
      logs.push(`[2/6] Created app registration — client ID: ${appRecord.appId}`)
    } else {
      logs.push(`[2/6] Found existing app registration — client ID: ${appRecord.appId}`)
      // Ensure the redirect URI is registered
      const existingSpas = appRecord.spa?.redirectUris ?? []
      if (!existingSpas.includes(effectiveRedirect)) {
        await graphRequest(accessToken, 'PATCH', `/applications/${appRecord.id}`, {
          spa: { redirectUris: [...existingSpas, effectiveRedirect] },
        })
        logs.push('[2/6] Added redirect URI to existing app registration.')
      }
    }

    // Step 3: Create service principal (enterprise app)
    logs.push('[3/6] Setting up Enterprise Application (service principal)...')
    let appServicePrincipal = (
      await graphRequest(
        accessToken,
        'GET',
        `/servicePrincipals?$filter=appId eq '${appRecord.appId}'&$select=id,appId,displayName`,
      )
    ).value?.[0]

    if (!appServicePrincipal) {
      appServicePrincipal = await graphRequest(accessToken, 'POST', '/servicePrincipals', {
        appId: appRecord.appId,
        tags: ['WindowsAzureActiveDirectoryIntegratedApp'],
      })
      logs.push(`[3/6] Created Enterprise Application — object ID: ${appServicePrincipal.id}`)
    } else {
      logs.push(`[3/6] Enterprise Application already exists — object ID: ${appServicePrincipal.id}`)
    }

    // Step 4: Resolve Microsoft Graph scope IDs
    logs.push('[4/6] Resolving Microsoft Graph permission scope IDs...')
    const graphServicePrincipal = (
      await graphRequest(
        accessToken,
        'GET',
        `/servicePrincipals?$filter=appId eq '${graphResourceAppId}'&$select=id,oauth2PermissionScopes`,
      )
    ).value?.[0]

    if (!graphServicePrincipal) {
      throw new Error('Microsoft Graph service principal not found in tenant.')
    }

    const resolvedScopes = []
    const unresolvedScopes = []
    for (const scopeValue of requiredGraphDelegatedScopes) {
      const scopeDefinition = (graphServicePrincipal.oauth2PermissionScopes ?? []).find(
        (scope) => scope.value === scopeValue,
      )
      if (scopeDefinition) {
        resolvedScopes.push({ id: scopeDefinition.id, type: 'Scope' })
      } else {
        unresolvedScopes.push(scopeValue)
      }
    }

    if (unresolvedScopes.length) {
      throw new Error(`Could not resolve Graph scopes: ${unresolvedScopes.join(', ')}`)
    }
    logs.push(`[4/6] Resolved ${resolvedScopes.length} permission scopes.`)

    // Step 5: Set required permissions on app registration (merge with existing)
    logs.push('[5/6] Updating app permissions...')
    const existingResources = appRecord.requiredResourceAccess ?? []
    const otherResources = existingResources.filter((r) => r.resourceAppId !== graphResourceAppId)
    await graphRequest(accessToken, 'PATCH', `/applications/${appRecord.id}`, {
      requiredResourceAccess: [
        ...otherResources,
        {
          resourceAppId: graphResourceAppId,
          resourceAccess: resolvedScopes,
        },
      ],
    })
    logs.push('[5/6] Updated app required permissions.')

    // Step 6: Admin consent grant
    logs.push('[6/6] Granting admin consent for delegated permissions...')
    const grantSearch = await graphRequest(
      accessToken,
      'GET',
      `/oauth2PermissionGrants?$filter=clientId eq '${appServicePrincipal.id}' and resourceId eq '${graphServicePrincipal.id}' and consentType eq 'AllPrincipals'`,
    )

    const existingGrant = (grantSearch.value ?? [])[0]
    if (existingGrant) {
      const mergedScopes = Array.from(
        new Set([...(existingGrant.scope ?? '').split(' ').filter(Boolean), ...requiredGraphDelegatedScopes]),
      ).join(' ')

      await graphRequest(accessToken, 'PATCH', `/oauth2PermissionGrants/${existingGrant.id}`, {
        scope: mergedScopes,
      })
      logs.push('[6/6] Updated admin consent grant.')
    } else {
      await graphRequest(accessToken, 'POST', '/oauth2PermissionGrants', {
        clientId: appServicePrincipal.id,
        consentType: 'AllPrincipals',
        resourceId: graphServicePrincipal.id,
        scope: requiredGraphDelegatedScopes.join(' '),
      })
      logs.push('[6/6] Created admin consent grant.')
    }

    const envValues = {
      VITE_AZURE_CLIENT_ID: appRecord.appId,
      VITE_AZURE_TENANT_ID: effectiveTenant,
    }
    updateEnvFile(envValues)

    logs.push('')
    logs.push('=== Onboarding complete ===')
    logs.push(`App name:   ${effectiveName}`)
    logs.push(`Client ID:  ${appRecord.appId}`)
    logs.push(`Tenant ID:  ${effectiveTenant}`)
    logs.push(`Action:     ${appCreated ? 'Created new app registration' : 'Updated existing app registration'}`)

    return {
      success: true,
      exitCode: 0,
      stdout: logs.join('\n'),
      stderr: '',
      error: '',
      envUpdated: true,
      envValues,
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error)
    logs.push('')
    logs.push(`=== Onboarding failed ===`)
    logs.push(errorMessage)
    return {
      success: false,
      exitCode: 1,
      stdout: logs.join('\n'),
      stderr: '',
      error: errorMessage,
      envUpdated: false,
      envValues: {},
    }
  } finally {
    if (mainWindowWasVisible && mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.restore()
      mainWindow.focus()
    }
  }
}

function startStaticServer() {
  return new Promise((resolve, reject) => {
    const staticApp = express()
    const distPath = path.join(__dirname, '..', 'dist')

    staticApp.use(express.static(distPath))
    staticApp.use((_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'))
    })

    const server = http.createServer(staticApp)
    server.on('error', reject)
    server.listen(desktopPort, '127.0.0.1', () => resolve(server))
  })
}

async function createMainWindow() {
  const preloadPath = path.join(__dirname, 'preload.cjs')

  mainWindow = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 1100,
    minHeight: 760,
    backgroundColor: '#050507',
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: preloadPath,
    },
  })

  // Allow the window to navigate to Microsoft login (redirect auth flow) and back.
  // Block any other external navigation that could happen accidentally.
  mainWindow.webContents.on('will-navigate', (event, url) => {
    const localBase = isDev
      ? process.env.VITE_DEV_SERVER_URL
      : `http://127.0.0.1:${desktopPort}`
    const isMicrosoftLogin =
      url.startsWith('https://login.microsoftonline.com') ||
      url.startsWith('https://login.microsoft.com')
    const isLocal = url.startsWith(localBase)
    if (!isLocal && !isMicrosoftLogin) {
      event.preventDefault()
    }
  })

  // MSAL Browser loginRedirect navigates the main window instead of using popups,
  // so we no longer need to allow about:blank windows.
  // All window.open() / target=_blank links open in the system browser.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url && url !== 'about:blank') {
      shell.openExternal(url)
    }
    return { action: 'deny' }
  })

  if (isDev) {
    await mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL)
    return
  }

  localServer = await startStaticServer()
  await mainWindow.loadURL(`http://127.0.0.1:${desktopPort}`)
}

app.whenReady().then(async () => {
  ipcMain.handle('runtime-config:get', async () => readRuntimeConfig())

  ipcMain.handle('config:reset', async () => {
    updateEnvFile({ VITE_AZURE_CLIENT_ID: '', VITE_AZURE_TENANT_ID: '' })
    return { success: true }
  })

  ipcMain.handle('tenant-setup:run', async (_event, input) => {
    if (!input || typeof input !== 'object') {
      return {
        success: false,
        exitCode: -1,
        stdout: '',
        stderr: '',
        error: 'No onboarding input was provided.',
        envUpdated: false,
        envValues: {},
      }
    }

    return runDirectOnboarding(input)
  })

  await createMainWindow()

  app.on('activate', async () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      await createMainWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('before-quit', () => {
  if (localServer) {
    localServer.close()
    localServer = null
  }
})
