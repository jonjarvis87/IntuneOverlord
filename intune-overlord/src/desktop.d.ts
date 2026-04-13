export {}

declare global {
  interface Window {
    intuneOverlordDesktop?: {
      isDesktop: boolean
      getRuntimeConfig?: () => Promise<{
        clientId: string
        tenantId: string
      }>
      resetConfig?: () => Promise<{ success: boolean }>
      runTenantSetup?: (input: {
        tenantId: string
        appName: string
        redirectUri: string
      }) => Promise<{
        success: boolean
        exitCode: number
        stdout: string
        stderr: string
        error: string
        envUpdated: boolean
        envValues: Record<string, string>
      }>
    }
  }
}
