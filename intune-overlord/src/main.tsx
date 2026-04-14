import { StrictMode, useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import './index.css'
import App from './App.tsx'

const pca = new PublicClientApplication({
  auth: {
    clientId: '3b96acc6-ee67-457f-84db-02a6baad96e7',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
  },
})

function RootApp() {
  const [ready, setReady] = useState(false)

  useEffect(() => {
    pca.initialize().then(() => setReady(true))
  }, [])

  if (!ready) return null

  return (
    <MsalProvider instance={pca}>
      <App />
    </MsalProvider>
  )
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <RootApp />
  </StrictMode>,
)
