import { StrictMode, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import './index.css'
import App from './App.tsx'

export type RuntimeAuthConfig = {
  clientId: string
  tenantId: string
}

const CONFIG_CLIENT_ID_KEY = 'intuneoverlord_client_id'
const CONFIG_TENANT_ID_KEY = 'intuneoverlord_tenant_id'

const loadConfig = (): RuntimeAuthConfig => ({
  clientId: localStorage.getItem(CONFIG_CLIENT_ID_KEY) || import.meta.env.VITE_AZURE_CLIENT_ID || '',
  tenantId: localStorage.getItem(CONFIG_TENANT_ID_KEY) || import.meta.env.VITE_AZURE_TENANT_ID || 'organizations',
})

const saveConfig = (config: RuntimeAuthConfig) => {
  if (config.clientId) {
    localStorage.setItem(CONFIG_CLIENT_ID_KEY, config.clientId)
    localStorage.setItem(CONFIG_TENANT_ID_KEY, config.tenantId)
  } else {
    localStorage.removeItem(CONFIG_CLIENT_ID_KEY)
    localStorage.removeItem(CONFIG_TENANT_ID_KEY)
  }
}

function RootApp() {
  const [authConfig, setAuthConfig] = useState<RuntimeAuthConfig>(loadConfig)
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(null)
  const pendingSignInRef = useRef(false)
  const prevConfigRef = useRef<RuntimeAuthConfig | null>(null)

  const initMsal = async (config: RuntimeAuthConfig) => {
    const pca = new PublicClientApplication({
      auth: {
        clientId: config.clientId || '00000000-0000-0000-0000-000000000000',
        authority: `https://login.microsoftonline.com/${config.tenantId || 'organizations'}`,
        redirectUri: window.location.origin,
        postLogoutRedirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: 'localStorage',
      },
    })

    await pca.initialize()

    if (
      pendingSignInRef.current &&
      config.clientId &&
      config.clientId !== '00000000-0000-0000-0000-000000000000' &&
      pca.getAllAccounts().length === 0
    ) {
      pendingSignInRef.current = false
      await pca.loginRedirect({
        scopes: ['DeviceManagementConfiguration.ReadWrite.All', 'DeviceManagementManagedDevices.Read.All', 'Group.Read.All'],
      })
      return
    }

    setMsalInstance(pca)
  }

  // Initialise MSAL on first render and whenever clientId/tenantId changes
  const configKey = `${authConfig.clientId}:${authConfig.tenantId}`
  const prevKey = prevConfigRef.current ? `${prevConfigRef.current.clientId}:${prevConfigRef.current.tenantId}` : null
  if (configKey !== prevKey) {
    prevConfigRef.current = authConfig
    void initMsal(authConfig)
  }

  if (!msalInstance) {
    return null
  }

  return (
    <MsalProvider instance={msalInstance}>
      <App
        authConfig={authConfig}
        onAuthConfigChange={(nextConfig, requestAutoSignIn) => {
          saveConfig(nextConfig)
          if (requestAutoSignIn) {
            pendingSignInRef.current = true
          }
          setMsalInstance(null)
          setAuthConfig(nextConfig)
        }}
      />
    </MsalProvider>
  )
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <RootApp />
  </StrictMode>,
)
