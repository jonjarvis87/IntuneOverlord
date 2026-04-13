# Intune Overlord

A desktop Electron app for bulk managing Microsoft Intune policy assignments via the Microsoft Graph API.

## Features

- **Bulk add** group assignments across multiple Intune policies at once
- **Bulk remove** group assignments across multiple Intune policies at once
- **Bulk delete** selected policies with a mandatory type-to-confirm safety modal
- **Delete ALL** — one-click button that selects every loaded policy and opens the delete confirmation modal (irreversible; requires typing `DELETE` to proceed)
- **Select all / Deselect all** — toggle button in the Policies panel to quickly select or clear all visible policies
- **View existing assignments** per policy with group display names resolved
- **Live group search** — search Entra ID groups by name directly in the UI
- **Export / Import** assignments as JSON or CSV
- **Policy categories** — Device Config, Compliance, Settings Catalog, Group Policy, Scripts, Remediations; collapsible with per-category checkbox selection
- **First-time setup** — automated Entra app registration, permissions, and admin consent via the built-in onboarding flow

## Bulk delete

> **Warning — this is irreversible.** Deleted policies cannot be recovered.

| Button | What it does |
|---|---|
| `⚠️ Delete policies` | Deletes the policies you have manually selected |
| `☢️ Delete ALL policies` | Selects **all** loaded policies and opens the confirm modal |

Both operations require you to type `DELETE` exactly in the confirmation dialog before the delete button is unlocked.

## Requirements

- Windows (Electron desktop app)
- Node.js 18+
- A Microsoft Entra (Azure AD) tenant with Intune

## Getting started

```cmd
Start-IntuneOverlord.cmd
```

This builds the app and launches it. On first run, enter your tenant ID and click **Create App** to complete automated onboarding.

To sign in manually with an existing app registration, add your `VITE_AZURE_CLIENT_ID` and `VITE_AZURE_TENANT_ID` to a `.env` file (see `.env.example`).

## Development

```bash
cd intune-overlord
npm install
npm run build
npm run desktop
```

## Required Graph API permissions

The app registration requires the following delegated permissions:

- `DeviceManagementConfiguration.ReadWrite.All`
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementScripts.ReadWrite.All`
- `Group.Read.All`

These are granted automatically during onboarding.
