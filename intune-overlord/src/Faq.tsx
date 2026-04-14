interface FaqProps {
  onBack: () => void
}

export function Faq({ onBack }: FaqProps) {
  return (
    <div className="faq-shell">
      <header className="faq-header">
        <div>
          <h1 className="faq-title">Documentation &amp; FAQ</h1>
          <p className="faq-subtitle">Everything you need to know about Intune Overlord</p>
        </div>
        <button type="button" className="secondary" onClick={onBack}>
          ← Back to app
        </button>
      </header>

      <div className="faq-grid">

        {/* ── Left column ── */}
        <div className="faq-col">

          <section className="faq-section">
            <h2>What is Intune Overlord?</h2>
            <p>
              Intune Overlord is a free web tool for IT administrators that lets you manage Microsoft
              Intune policy assignments in bulk. Instead of clicking through the Intune portal
              one policy at a time, you can add or remove group assignments across dozens (or
              hundreds) of policies simultaneously.
            </p>
            <p>
              It was built by <a href="https://www.cloudendpoint.ai" target="_blank" rel="noreferrer">Jon Jarvis</a>,
              Microsoft MVP in Intune, to solve a real-world problem faced by engineers managing
              large Intune estates.
            </p>
          </section>

          <section className="faq-section">
            <h2>Getting started</h2>
            <ol className="faq-ol">
              <li>
                Navigate to the app and click <strong>Sign in</strong>. You'll be redirected to
                Microsoft to authenticate with your work account.
              </li>
              <li>
                If this is the first sign-in from your organisation, Microsoft will show a
                <strong> consent screen</strong> listing the permissions the app requires.
                An administrator must approve these on behalf of the organisation.
              </li>
              <li>
                Once signed in, your Intune policies load automatically. Select the policies
                you want to work with on the left, build your assignment list on the right,
                then hit <strong>Bulk add</strong> or <strong>Bulk remove</strong>.
              </li>
            </ol>
          </section>

          <section className="faq-section">
            <h2>What permissions does it need?</h2>
            <p>The app requests four delegated Microsoft Graph permissions:</p>
            <ul className="faq-ul">
              <li>
                <code>DeviceManagementConfiguration.ReadWrite.All</code> — read and write
                device configuration and Settings Catalog policies
              </li>
              <li>
                <code>DeviceManagementManagedDevices.Read.All</code> — read compliance
                policy data
              </li>
              <li>
                <code>DeviceManagementScripts.ReadWrite.All</code> — read and write
                PowerShell scripts and Remediations
              </li>
              <li>
                <code>Group.Read.All</code> — search for and resolve Entra ID group names
              </li>
            </ul>
            <p>
              These are <strong>delegated</strong> permissions — the app acts as the signed-in
              user and can only do what that user is already permitted to do in Intune.
              No application-level (background) access is used.
            </p>
          </section>

          <section className="faq-section">
            <h2>Supported policy types</h2>
            <ul className="faq-ul">
              <li><strong>Device Config</strong> — classic device configuration profiles</li>
              <li><strong>Compliance</strong> — device compliance policies</li>
              <li><strong>Settings Catalog</strong> — modern settings catalog profiles</li>
              <li><strong>Group Policy</strong> — ADMX-backed administrative templates</li>
              <li><strong>Scripts</strong> — PowerShell scripts</li>
              <li><strong>Remediations</strong> — proactive remediation scripts</li>
            </ul>
            <p>
              All six types are loaded in parallel when you sign in. If your account
              doesn't have access to a particular type, it is silently skipped.
            </p>
          </section>

          <section className="faq-section">
            <h2>Is my data stored anywhere?</h2>
            <p>
              No. Intune Overlord is a pure browser-based application. All communication
              goes directly between your browser and the Microsoft Graph API — no data
              passes through any intermediate server. Nothing is logged or stored outside
              of your own browser session.
            </p>
            <p>
              Your Microsoft authentication token is held in browser localStorage by the
              MSAL library (standard Microsoft behaviour) and is scoped to your session.
            </p>
          </section>

        </div>

        {/* ── Right column ── */}
        <div className="faq-col">

          <section className="faq-section">
            <h2>Bulk add assignments</h2>
            <p>
              Select one or more policies on the left panel. In the right panel (Assign tab),
              add the groups you want to assign using the group search picker. Choose
              <strong> Include</strong> or <strong>Exclude</strong> for each group, then
              click <strong>Bulk add assignments</strong>.
            </p>
            <p>
              The tool fetches the existing assignments for every selected policy in parallel,
              merges your new groups in, deduplicates, and pushes the updated list back —
              so existing assignments are never accidentally removed.
            </p>
          </section>

          <section className="faq-section">
            <h2>Bulk remove assignments</h2>
            <p>
              Works the same way as bulk add, but in reverse. Add the groups you want to
              remove in the right panel, then click <strong>Bulk remove assignments</strong>.
              The tool fetches each policy's current assignments, strips out the matching
              groups, and pushes the result back.
            </p>
            <p>
              You can remove a group regardless of whether it was set to Include or Exclude —
              the match is on group ID only.
            </p>
          </section>

          <section className="faq-section">
            <h2>Exporting and importing assignments</h2>
            <p>
              <strong>Export</strong> — select policies and click Export in the Assign tab.
              This downloads a JSON file containing every group assignment across all selected
              policies, including policy name, type, group ID and include/exclude mode.
            </p>
            <p>
              <strong>Import</strong> — click Import and select a <code>.json</code> or
              <code>.csv</code> file. The file should contain <code>groupId</code> and
              <code>mode</code> columns/keys. The imported rows populate the assignment
              builder so you can apply them to any selected policies.
            </p>
          </section>

          <section className="faq-section">
            <h2>Viewing current assignments</h2>
            <p>
              Select one or more policies and click the <strong>Current</strong> tab in the
              right panel. The app fetches all assignments and resolves group names in a
              single batch call, showing each group's display name alongside its object ID
              and include/exclude status.
            </p>
          </section>

          <section className="faq-section">
            <h2>Deleting policies</h2>
            <p>
              Select the policies you want to remove and click <strong>⚠️ Delete policies</strong>.
              A confirmation dialog will appear requiring you to type <code>DELETE</code> before
              the action proceeds. Deletion is permanent and cannot be undone.
            </p>
            <p>
              <strong>☢️ Delete ALL policies</strong> selects every loaded policy and opens
              the same confirmation dialog. Use with extreme caution.
            </p>
          </section>

          <section className="faq-section">
            <h2>Why aren't all my policies showing?</h2>
            <p>
              The Microsoft Graph API paginates results. Intune Overlord requests up to 999
              items per page and follows pagination links automatically, so you should see
              all policies in most cases.
            </p>
            <p>
              In rare cases the Intune backend returns a server error on a pagination token
              (a known Microsoft bug, most common on Settings Catalog). When this happens
              the app returns the policies fetched so far rather than failing entirely.
              Refreshing the policy list may return a different — sometimes fuller — result.
            </p>
          </section>

          <section className="faq-section">
            <h2>Who can use it?</h2>
            <p>
              Anyone with a Microsoft work or school account that has Intune administrator
              access. The app is multi-tenant — admins from any organisation can sign in.
              The first person from a new tenant will be prompted to grant admin consent
              for the required permissions.
            </p>
          </section>

        </div>
      </div>
    </div>
  )
}
