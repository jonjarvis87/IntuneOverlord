import { StrictMode, useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import './index.css'
import App from './App.tsx'

export type RuntimeAuthConfig = {
  clientId: string
  tenantId: string
}

const defaultAuthConfig: RuntimeAuthConfig = {
  clientId: import.meta.env.VITE_AZURE_CLIENT_ID || '',
  tenantId: import.meta.env.VITE_AZURE_TENANT_ID || 'organizations',
}

const desktopRedirectUri = window.intuneOverlordDesktop?.isDesktop ? 'http://127.0.0.1:4783' : window.location.origin

function RootApp() {
  const [authConfig, setAuthConfig] = useState<RuntimeAuthConfig>(defaultAuthConfig)
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(null)
  // Use a ref (not state) so the initMsal effect always reads the latest value
  // without needing it in the dependency array — avoids stale closure issues.
  const pendingSignInRef = useRef(false)

  // Load live config from .env on startup (desktop only)
  useEffect(() => {
    const loadRuntimeConfig = async () => {
      const runtimeConfig = await window.intuneOverlordDesktop?.getRuntimeConfig?.()

      if (!runtimeConfig) {
        return
      }

      setAuthConfig({
        clientId: runtimeConfig.clientId || defaultAuthConfig.clientId,
        tenantId: runtimeConfig.tenantId || defaultAuthConfig.tenantId,
      })
    }

    void loadRuntimeConfig()
  }, [])

  // Re-create + initialise MSAL whenever the client ID or tenant changes.
  // initialize() must be awaited before passing to MsalProvider so that
  // handleRedirectPromise() runs and the redirect token is processed.
  useEffect(() => {
    let cancelled = false

    const initMsal = async () => {
      const pca = new PublicClientApplication({
        auth: {
          clientId: authConfig.clientId || '00000000-0000-0000-0000-000000000000',
          authority: `https://login.microsoftonline.com/${authConfig.tenantId || 'organizations'}`,
          redirectUri: desktopRedirectUri,
          postLogoutRedirectUri: desktopRedirectUri,
        },
        cache: {
          cacheLocation: 'localStorage',
        },
      })

      await pca.initialize()

      if (cancelled) return

      // If onboarding just completed, fire loginRedirect immediately with this
      // correctly-configured PCA — before it's handed to MsalProvider.
      if (
        pendingSignInRef.current &&
        authConfig.clientId &&
        authConfig.clientId !== '00000000-0000-0000-0000-000000000000' &&
        pca.getAllAccounts().length === 0
      ) {
        pendingSignInRef.current = false
        await pca.loginRedirect({
          scopes: ['DeviceManagementConfiguration.ReadWrite.All', 'DeviceManagementManagedDevices.Read.All', 'Group.Read.All'],
        })
        // loginRedirect navigates the window away — nothing below runs.
        return
      }

      setMsalInstance(pca)
    }

    void initMsal()
    return () => {
      cancelled = true
    }
  }, [authConfig.clientId, authConfig.tenantId])

  if (!msalInstance) {
    return null
  }

  return (
    <MsalProvider instance={msalInstance}>
      <App
        authConfig={authConfig}
        onAuthConfigChange={(nextConfig, requestAutoSignIn) => {
          if (requestAutoSignIn) {
            pendingSignInRef.current = true
          }
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
