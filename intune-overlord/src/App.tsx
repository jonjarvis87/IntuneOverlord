import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent } from 'react'
import Papa from 'papaparse'
import { useIsAuthenticated, useMsal } from '@azure/msal-react'
import type { RuntimeAuthConfig } from './main'
import './App.css'

type AssignmentMode = 'include' | 'exclude'

type GroupResult = {
  id: string
  displayName: string
}

// ---------------------------------------------------------------------------
// Confirm Delete Modal
// ---------------------------------------------------------------------------
interface ConfirmDeleteModalProps {
  policies: PolicyRecord[]
  onConfirm: () => void
  onCancel: () => void
}

function ConfirmDeleteModal({ policies, onConfirm, onCancel }: ConfirmDeleteModalProps) {
  const [typed, setTyped] = useState('')
  const confirmed = typed === 'DELETE'

  return (
    <div className="modal-backdrop">
      <div className="modal modal--danger">
        <div className="modal-warning-banner">
          <span className="modal-warning-icon">⚠️</span>
          <span>PERMANENT DELETION — THIS CANNOT BE UNDONE</span>
        </div>
        <div className="modal-body">
          <p className="modal-lead">
            You are about to <strong>permanently delete {policies.length} {policies.length === 1 ? 'policy' : 'policies'}</strong> from Intune.
            This will remove the policy configuration entirely from your tenant — not just its assignments.
          </p>
          <ul className="modal-policy-list">
            {policies.slice(0, 20).map((p) => (
              <li key={p.id}>
                <span className="modal-policy-kind">{p.kind}</span>
                <span className="modal-policy-name">{p.displayName}</span>
              </li>
            ))}
            {policies.length > 20 && (
              <li className="modal-policy-overflow">…and {policies.length - 20} more</li>
            )}
          </ul>
          <p className="modal-confirm-label">
            Type <strong>DELETE</strong> to confirm:
          </p>
          <input
            className="modal-confirm-input"
            value={typed}
            onChange={(e) => setTyped(e.target.value)}
            placeholder="DELETE"
            autoFocus
          />
        </div>
        <div className="modal-footer">
          <button type="button" className="secondary" onClick={onCancel}>
            Cancel
          </button>
          <button
            type="button"
            className="danger modal-delete-btn"
            disabled={!confirmed}
            onClick={onConfirm}
          >
            Delete {policies.length} {policies.length === 1 ? 'policy' : 'policies'}
          </button>
        </div>
      </div>
    </div>
  )
}

interface GroupPickerProps {
  resolvedName: string
  groupId: string
  disabled: boolean
  onSelect: (id: string, name: string) => void
  getToken: () => Promise<string>
}

function GroupPicker({ resolvedName, groupId, disabled, onSelect, getToken }: GroupPickerProps) {
  const [query, setQuery] = useState(resolvedName || groupId)
  const [results, setResults] = useState<GroupResult[]>([])
  const [isSearching, setIsSearching] = useState(false)
  const [isOpen, setIsOpen] = useState(false)
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const wrapperRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    setQuery(resolvedName || groupId)
  }, [resolvedName, groupId])

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) {
        setIsOpen(false)
      }
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  const handleQueryChange = (q: string) => {
    setQuery(q)
    if (debounceRef.current) clearTimeout(debounceRef.current)
    if (!q.trim() || q.trim().length < 2) {
      setResults([])
      setIsOpen(false)
      return
    }
    debounceRef.current = setTimeout(() => {
      void (async () => {
        setIsSearching(true)
        try {
          const token = await getToken()
          const resp = await fetch(
            `https://graph.microsoft.com/beta/groups?$search="displayName:${encodeURIComponent(q)}"&$select=id,displayName&$top=12&$orderby=displayName`,
            {
              headers: {
                Authorization: `Bearer ${token}`,
                ConsistencyLevel: 'eventual',
              },
            },
          )
          if (resp.ok) {
            const data = (await resp.json()) as { value?: GroupResult[] }
            setResults(data.value ?? [])
            setIsOpen(true)
          }
        } catch {
          // ignore search errors
        } finally {
          setIsSearching(false)
        }
      })()
    }, 350)
  }

  const select = (group: GroupResult) => {
    setQuery(group.displayName)
    setResults([])
    setIsOpen(false)
    onSelect(group.id, group.displayName)
  }

  const handleBlur = () => {
    // Close dropdown
    setIsOpen(false)
    const trimmed = query.trim()
    // If nothing selected from dropdown (resolvedName not matching query),
    // commit whatever is in the input as the raw groupId so manual GUIDs work.
    if (trimmed && trimmed !== resolvedName) {
      onSelect(trimmed, '')
    }
  }

  return (
    <div className="group-picker" ref={wrapperRef}>
      <input
        className="group-picker-input"
        value={query}
        placeholder="Search or paste Group Object ID…"
        disabled={disabled}
        onChange={(e) => handleQueryChange(e.target.value)}
        onFocus={() => results.length > 0 && setIsOpen(true)}
        onBlur={handleBlur}
      />
      {resolvedName && groupId && groupId !== resolvedName && (
        <span className="group-picker-id" title={groupId}>{groupId}</span>
      )}
      {!resolvedName && groupId && (
        <span className="group-picker-id" title={groupId}>{groupId}</span>
      )}
      {isSearching && <span className="group-picker-hint">Searching…</span>}
      {isOpen && results.length > 0 && (
        <ul className="group-picker-dropdown">
          {results.map((g) => (
            <li key={g.id} className="group-picker-option" onMouseDown={() => select(g)}>
              <span className="group-picker-option-name">{g.displayName}</span>
              <span className="group-picker-option-id">{g.id}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

type AssignmentDraft = {
  groupId: string
  mode: AssignmentMode
}

type PolicyRecord = {
  id: string
  displayName: string
  kind: 'Device Config' | 'Compliance' | 'Settings Catalog' | 'Group Policy' | 'Scripts' | 'Remediations'
  resourcePath: string
}

const graphScopes = [
  'DeviceManagementConfiguration.ReadWrite.All',
  'DeviceManagementManagedDevices.Read.All',
  'DeviceManagementScripts.ReadWrite.All',
  'Group.Read.All',
]

const graphApiVersion = 'beta'

const normalizeMode = (value?: string): AssignmentMode => {
  if (value?.toLowerCase() === 'exclude') {
    return 'exclude'
  }

  return 'include'
}

const buildAssignmentPayload = (draft: AssignmentDraft) => {
  const targetType =
    draft.mode === 'exclude'
      ? '#microsoft.graph.exclusionGroupAssignmentTarget'
      : '#microsoft.graph.groupAssignmentTarget'

  return {
    target: {
      '@odata.type': targetType,
      groupId: draft.groupId,
    },
  }
}

const parseAssignmentFromGraph = (assignment: Record<string, unknown>) => {
  const target = assignment.target as Record<string, unknown> | undefined

  if (!target || typeof target.groupId !== 'string') {
    return null
  }

  const targetType = typeof target['@odata.type'] === 'string' ? target['@odata.type'] : ''
  const mode: AssignmentMode = targetType.toLowerCase().includes('exclusion') ? 'exclude' : 'include'

  return {
    groupId: target.groupId,
    mode,
  } satisfies AssignmentDraft
}

const trimAndDedupeDrafts = (drafts: AssignmentDraft[]) => {
  const unique = new Map<string, AssignmentDraft>()

  for (const draft of drafts) {
    const cleanedId = draft.groupId.trim()

    if (!cleanedId) {
      continue
    }

    const normalized: AssignmentDraft = {
      groupId: cleanedId,
      mode: normalizeMode(draft.mode),
    }

    unique.set(`${normalized.groupId}:${normalized.mode}`, normalized)
  }

  return Array.from(unique.values())
}

type AppProps = {
  authConfig: RuntimeAuthConfig
  onAuthConfigChange: (nextConfig: RuntimeAuthConfig, requestAutoSignIn?: boolean) => void
}

// ---------------------------------------------------------------------------
// Module-level constants and pure helpers (stable across renders)
// ---------------------------------------------------------------------------

const categoryOrder: PolicyRecord['kind'][] = [
  'Device Config',
  'Compliance',
  'Settings Catalog',
  'Group Policy',
  'Scripts',
  'Remediations',
]

type ExistingAssignmentRow = {
  groupId: string
  groupName: string
  mode: AssignmentMode
}

type ExistingPolicyBlock = {
  policyId: string
  policyName: string
  policyKind: PolicyRecord['kind']
  assignments: ExistingAssignmentRow[]
}

const downloadText = (fileName: string, content: string, mimeType: string) => {
  const blob = new Blob([content], { type: mimeType })
  const objectUrl = URL.createObjectURL(blob)
  const linkElement = document.createElement('a')
  linkElement.href = objectUrl
  linkElement.download = fileName
  linkElement.click()
  URL.revokeObjectURL(objectUrl)
}

function App({ authConfig, onAuthConfigChange }: AppProps) {
  const { instance, accounts } = useMsal()
  const isAuthenticated = useIsAuthenticated()

  const [isBusy, setIsBusy] = useState(false)
  const [policies, setPolicies] = useState<PolicyRecord[]>([])
  const [selectedPolicyIds, setSelectedPolicyIds] = useState<Set<string>>(new Set())
  const [collapsedCategories, setCollapsedCategories] = useState<Set<string>>(new Set())
  const [policySearch, setPolicySearch] = useState('')
  const [draftAssignments, setDraftAssignments] = useState<AssignmentDraft[]>([
    { groupId: '', mode: 'include' },
  ])
  const [draftGroupNames, setDraftGroupNames] = useState<string[]>([''])
  const [activeTab, setActiveTab] = useState<'assign' | 'view'>('assign')
  const [existingAssignments, setExistingAssignments] = useState<ExistingPolicyBlock[]>([])
  const [isViewLoading, setIsViewLoading] = useState(false)
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false)
  const [statusMessage, setStatusMessage] = useState('Sign in and load policies to begin.')
  const [setupClientId, setSetupClientId] = useState('')
  const [setupTenantId, setSetupTenantId] = useState('')

  const clientIdConfigured = Boolean(authConfig.clientId)

  // Auto-load policies as soon as the user successfully signs in.
  useEffect(() => {
    if (isAuthenticated && policies.length === 0 && !isBusy) {
      void loadPolicies()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAuthenticated])

  const filteredPolicies = useMemo(() => {
    const searchValue = policySearch.trim().toLowerCase()

    if (!searchValue) {
      return policies
    }

    return policies.filter(
      (policy) =>
        policy.displayName.toLowerCase().includes(searchValue) ||
        policy.kind.toLowerCase().includes(searchValue),
    )
  }, [policies, policySearch])

  const groupedPolicies = useMemo(() => {
    const groups = new Map<PolicyRecord['kind'], PolicyRecord[]>()
    for (const kind of categoryOrder) {
      groups.set(kind, [])
    }
    for (const policy of filteredPolicies) {
      groups.get(policy.kind)?.push(policy)
    }
    return groups
  }, [filteredPolicies])

  const toggleCategory = (kind: string) => {
    setCollapsedCategories((current) => {
      const next = new Set(current)
      if (next.has(kind)) {
        next.delete(kind)
      } else {
        next.add(kind)
      }
      return next
    })
  }

  const toggleCategorySelection = (kind: PolicyRecord['kind']) => {
    const group = groupedPolicies.get(kind) ?? []
    const allSelected = group.length > 0 && group.every((p) => selectedPolicyIds.has(p.id))
    setSelectedPolicyIds((current) => {
      const next = new Set(current)
      for (const policy of group) {
        if (allSelected) {
          next.delete(policy.id)
        } else {
          next.add(policy.id)
        }
      }
      return next
    })
  }

  const allFilteredSelected = useMemo(
    () => filteredPolicies.length > 0 && filteredPolicies.every((p) => selectedPolicyIds.has(p.id)),
    [filteredPolicies, selectedPolicyIds],
  )

  const handleSelectAll = () => {
    if (allFilteredSelected) {
      setSelectedPolicyIds(new Set())
    } else {
      setSelectedPolicyIds(new Set(filteredPolicies.map((p) => p.id)))
    }
  }

  const handleDeleteAll = () => {
    // Select every loaded policy then open the confirm modal in one batched render
    setSelectedPolicyIds(new Set(policies.map((p) => p.id)))
    setDeleteConfirmOpen(true)
  }

  const selectedPolicies = useMemo(
    () => policies.filter((policy) => selectedPolicyIds.has(policy.id)),
    [policies, selectedPolicyIds],
  )

  const getToken = useCallback(async () => {
    const account = accounts[0]

    if (!account) {
      throw new Error('No signed-in account found.')
    }

    try {
      const result = await instance.acquireTokenSilent({
        account,
        scopes: graphScopes,
      })

      return result.accessToken
    } catch {
      // Silent refresh failed — redirect for interactive re-auth.
      // The page will navigate away; execution does not continue past this.
      await instance.acquireTokenRedirect({
        account,
        scopes: graphScopes,
      })
      throw new Error('Redirecting for re-authentication...')
    }
  }, [accounts, instance])

  const graphRequest = async <T,>(
    token: string,
    endpoint: string,
    method: 'GET' | 'POST' | 'DELETE' = 'GET',
    body?: unknown,
  ) => {
    const response = await fetch(`https://graph.microsoft.com/${graphApiVersion}${endpoint}`, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: body ? JSON.stringify(body) : undefined,
    })

    const text = await response.text()

    if (!response.ok) {
      throw new Error(`Graph ${method} ${endpoint} failed: ${response.status} ${text}`)
    }

    if (!text.trim()) {
      return undefined as T
    }

    return JSON.parse(text) as T
  }

  const fetchAssignments = async (token: string, policy: PolicyRecord) => {
    const result = await graphRequest<{ value?: Record<string, unknown>[] }>(
      token,
      `${policy.resourcePath}/${policy.id}/assignments`,
    )

    const assignments = (result.value ?? [])
      .map((assignment) => parseAssignmentFromGraph(assignment))
      .filter((assignment): assignment is AssignmentDraft => assignment !== null)

    return trimAndDedupeDrafts(assignments)
  }

  const assignPolicy = async (token: string, policy: PolicyRecord, assignments: AssignmentDraft[]) => {
    await graphRequest(
      token,
      `${policy.resourcePath}/${policy.id}/assign`,
      'POST',
      {
        assignments: assignments.map((assignment) => buildAssignmentPayload(assignment)),
      },
    )
  }

  const loadPolicies = async () => {
    setIsBusy(true)
    setStatusMessage('Loading Intune policies...')

    try {
      const token = await getToken()
      const safeGet = async (endpoint: string) => {
        try {
          return await graphRequest<{ value?: Record<string, unknown>[] }>(token, endpoint)
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err)
          // 403 = missing scope; 404 = not licensed — skip silently
          if (msg.includes('403') || msg.includes('404')) return { value: [] }
          throw err
        }
      }

      const [deviceConfigs, compliancePolicies, configurationPolicies, groupPolicyConfigs, platformScripts, remediations] = await Promise.all([
        safeGet('/deviceManagement/deviceConfigurations?$select=id,displayName'),
        safeGet('/deviceManagement/deviceCompliancePolicies?$select=id,displayName'),
        safeGet('/deviceManagement/configurationPolicies?$select=id,name'),
        safeGet('/deviceManagement/groupPolicyConfigurations?$select=id,displayName'),
        safeGet('/deviceManagement/deviceManagementScripts?$select=id,displayName'),
        safeGet('/deviceManagement/deviceHealthScripts?$select=id,displayName'),
      ])

      const mappedPolicies: PolicyRecord[] = [
        ...(deviceConfigs.value ?? []).flatMap((policy) => {
          if (typeof policy.id !== 'string' || typeof policy.displayName !== 'string') {
            return []
          }

          return [
            {
              id: policy.id,
              displayName: policy.displayName,
              kind: 'Device Config',
              resourcePath: '/deviceManagement/deviceConfigurations',
            } satisfies PolicyRecord,
          ]
        }),
        ...(compliancePolicies.value ?? []).flatMap((policy) => {
          if (typeof policy.id !== 'string' || typeof policy.displayName !== 'string') {
            return []
          }

          return [
            {
              id: policy.id,
              displayName: policy.displayName,
              kind: 'Compliance',
              resourcePath: '/deviceManagement/deviceCompliancePolicies',
            } satisfies PolicyRecord,
          ]
        }),
        ...(configurationPolicies.value ?? []).flatMap((policy) => {
          if (typeof policy.id !== 'string' || typeof policy.name !== 'string') {
            return []
          }

          return [
            {
              id: policy.id,
              displayName: policy.name,
              kind: 'Settings Catalog',
              resourcePath: '/deviceManagement/configurationPolicies',
            } satisfies PolicyRecord,
          ]
        }),
        ...(groupPolicyConfigs.value ?? []).flatMap((policy) => {
          if (typeof policy.id !== 'string' || typeof policy.displayName !== 'string') {
            return []
          }

          return [
            {
              id: policy.id,
              displayName: policy.displayName,
              kind: 'Group Policy',
              resourcePath: '/deviceManagement/groupPolicyConfigurations',
            } satisfies PolicyRecord,
          ]
        }),
        ...(platformScripts.value ?? []).flatMap((policy) => {
          if (typeof policy.id !== 'string' || typeof policy.displayName !== 'string') {
            return []
          }

          return [
            {
              id: policy.id,
              displayName: policy.displayName,
              kind: 'Scripts',
              resourcePath: '/deviceManagement/deviceManagementScripts',
            } satisfies PolicyRecord,
          ]
        }),
        ...(remediations.value ?? []).flatMap((policy) => {
          if (typeof policy.id !== 'string' || typeof policy.displayName !== 'string') {
            return []
          }

          return [
            {
              id: policy.id,
              displayName: policy.displayName,
              kind: 'Remediations',
              resourcePath: '/deviceManagement/deviceHealthScripts',
            } satisfies PolicyRecord,
          ]
        }),
      ].sort((left, right) => left.displayName.localeCompare(right.displayName))

      setPolicies(mappedPolicies)
      setSelectedPolicyIds(new Set())
      setStatusMessage(`Loaded ${mappedPolicies.length} policies.`)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to load policies.'
      setStatusMessage(message)
    } finally {
      setIsBusy(false)
    }
  }

  const updateDraft = (index: number, next: AssignmentDraft) => {
    setDraftAssignments((current) => current.map((draft, draftIndex) => (draftIndex === index ? next : draft)))
  }

  const appendDraft = () => {
    setDraftAssignments((current) => [...current, { groupId: '', mode: 'include' }])
    setDraftGroupNames((current) => [...current, ''])
  }

  const removeDraft = (index: number) => {
    setDraftAssignments((current) => {
      if (current.length === 1) {
        return [{ groupId: '', mode: 'include' }]
      }

      return current.filter((_, draftIndex) => draftIndex !== index)
    })
    setDraftGroupNames((current) => {
      if (current.length === 1) return ['']
      return current.filter((_, draftIndex) => draftIndex !== index)
    })
  }

  const togglePolicy = (policyId: string, checked: boolean) => {
    setSelectedPolicyIds((current) => {
      const next = new Set(current)

      if (checked) {
        next.add(policyId)
      } else {
        next.delete(policyId)
      }

      return next
    })
  }

  const loadExistingAssignments = async () => {
    if (!selectedPolicies.length) {
      setExistingAssignments([])
      return
    }
    setIsViewLoading(true)
    try {
      const token = await getToken()
      const blocks: ExistingPolicyBlock[] = []
      const allGroupIds = new Set<string>()

      // Fetch assignments for every selected policy in parallel
      const fetched = await Promise.all(
        selectedPolicies.map(async (policy) => {
          const assignments = await fetchAssignments(token, policy)
          for (const a of assignments) allGroupIds.add(a.groupId)
          return { policy, assignments }
        }),
      )

      // Resolve all group IDs to display names in one batch call
      const groupNameMap = new Map<string, string>()
      const idsArray = Array.from(allGroupIds)
      if (idsArray.length) {
        try {
          const resp = await fetch('https://graph.microsoft.com/v1.0/directoryObjects/getByIds', {
            method: 'POST',
            headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: idsArray, types: ['group'] }),
          })
          if (resp.ok) {
            const data = (await resp.json()) as { value?: Array<{ id: string; displayName?: string }> }
            for (const obj of data.value ?? []) {
              if (obj.id && obj.displayName) groupNameMap.set(obj.id, obj.displayName)
            }
          }
        } catch {
          // name resolution is best-effort; fall back to IDs
        }
      }

      for (const { policy, assignments } of fetched) {
        blocks.push({
          policyId: policy.id,
          policyName: policy.displayName,
          policyKind: policy.kind,
          assignments: assignments.map((a) => ({
            groupId: a.groupId,
            groupName: groupNameMap.get(a.groupId) ?? '',
            mode: a.mode,
          })),
        })
      }
      setExistingAssignments(blocks)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to load assignments.'
      setStatusMessage(message)
    } finally {
      setIsViewLoading(false)
    }
  }

  const switchToViewTab = () => {
    setActiveTab('view')
    void loadExistingAssignments()
  }

  const runBulkDelete = async () => {
    setDeleteConfirmOpen(false)
    setIsBusy(true)
    setStatusMessage(`Deleting ${selectedPolicies.length} ${selectedPolicies.length === 1 ? 'policy' : 'policies'}…`)
    try {
      const token = await getToken()
      // Fire all deletes in parallel; collect individual pass/fail
      const outcomes = await Promise.allSettled(
        selectedPolicies.map((policy) =>
          graphRequest(token, `${policy.resourcePath}/${policy.id}`, 'DELETE').then(() => policy.id)
        )
      )
      const succeededIds = new Set(
        outcomes
          .filter((r): r is PromiseFulfilledResult<string> => r.status === 'fulfilled')
          .map((r) => r.value)
      )
      outcomes.forEach((r, i) => {
        if (r.status === 'rejected') {
          console.error(`Failed to delete ${selectedPolicies[i].displayName}:`, r.reason)
        }
      })
      const failed = outcomes.length - succeededIds.size
      // Only remove policies that were actually deleted
      setPolicies((current) => current.filter((p) => !succeededIds.has(p.id)))
      setSelectedPolicyIds(new Set())
      setStatusMessage(
        `Deleted ${succeededIds.size} ${succeededIds.size === 1 ? 'policy' : 'policies'}${
          failed ? `, ${failed} failed — check console for details` : '.'
        }`
      )
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Bulk delete failed.'
      setStatusMessage(message)
    } finally {
      setIsBusy(false)
    }
  }

  const runBulkAdd = async () => {
    const additions = trimAndDedupeDrafts(draftAssignments)

    if (!selectedPolicies.length || !additions.length) {
      setStatusMessage('Select at least one policy and one valid assignment row.')
      return
    }

    setIsBusy(true)
    setStatusMessage(`Adding assignments to ${selectedPolicies.length} policies...`)

    try {
      const token = await getToken()
      // Fetch all existing assignments in parallel, then push all updates in parallel
      const existing = await Promise.all(selectedPolicies.map((policy) => fetchAssignments(token, policy)))
      await Promise.all(
        selectedPolicies.map((policy, i) =>
          assignPolicy(token, policy, trimAndDedupeDrafts([...existing[i], ...additions]))
        )
      )
      setStatusMessage(`Added assignments to ${selectedPolicies.length} policies.`)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Bulk add failed.'
      setStatusMessage(message)
    } finally {
      setIsBusy(false)
    }
  }

  const runBulkRemove = async () => {
    const removals = trimAndDedupeDrafts(draftAssignments).map((draft) => draft.groupId)

    if (!selectedPolicies.length || !removals.length) {
      setStatusMessage('Select at least one policy and one group ID to remove.')
      return
    }

    setIsBusy(true)
    setStatusMessage(`Removing assignments from ${selectedPolicies.length} policies...`)

    try {
      const token = await getToken()
      // Fetch all current assignments in parallel, then push all filtered updates in parallel
      const existing = await Promise.all(selectedPolicies.map((policy) => fetchAssignments(token, policy)))
      await Promise.all(
        selectedPolicies.map((policy, i) =>
          assignPolicy(token, policy, existing[i].filter((a) => !removals.includes(a.groupId)))
        )
      )
      setStatusMessage(`Removed assignments from ${selectedPolicies.length} policies.`)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Bulk remove failed.'
      setStatusMessage(message)
    } finally {
      setIsBusy(false)
    }
  }

  const exportAssignments = async () => {
    if (!selectedPolicies.length) {
      setStatusMessage('Select at least one policy to export assignments.')
      return
    }

    setIsBusy(true)
    setStatusMessage(`Exporting assignments from ${selectedPolicies.length} policies...`)

    try {
      const token = await getToken()
      // Fetch all assignments in parallel
      const perPolicy = await Promise.all(
        selectedPolicies.map((policy) =>
          fetchAssignments(token, policy).then((assignments) => ({ policy, assignments }))
        )
      )
      const exportRows = perPolicy.flatMap(({ policy, assignments }) =>
        assignments.map((assignment) => ({
          policyId: policy.id,
          policyName: policy.displayName,
          policyType: policy.kind,
          groupId: assignment.groupId,
          mode: assignment.mode,
        }))
      )
      downloadText(
        'intune-overlord-assignments.json',
        JSON.stringify(exportRows, null, 2),
        'application/json',
      )
      setStatusMessage(`Exported ${exportRows.length} assignments.`)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Export failed.'
      setStatusMessage(message)
    } finally {
      setIsBusy(false)
    }
  }

  const importAssignments = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]

    if (!file) {
      return
    }

    try {
      let importedDrafts: AssignmentDraft[] = []

      if (file.name.toLowerCase().endsWith('.json')) {
        const text = await file.text()
        const parsed = JSON.parse(text) as Array<{ groupId?: string; mode?: string }>

        importedDrafts = (parsed ?? []).flatMap((item) => {
          if (typeof item.groupId !== 'string') {
            return []
          }

          return [
            {
              groupId: item.groupId,
              mode: normalizeMode(item.mode),
            } satisfies AssignmentDraft,
          ]
        })
      } else {
        const text = await file.text()
        const parsed = Papa.parse<{ groupId?: string; mode?: string }>(text, {
          header: true,
          skipEmptyLines: true,
        })

        importedDrafts = parsed.data.flatMap((row) => {
          if (typeof row.groupId !== 'string') {
            return []
          }

          return [
            {
              groupId: row.groupId,
              mode: normalizeMode(row.mode),
            } satisfies AssignmentDraft,
          ]
        })
      }

      const cleaned = trimAndDedupeDrafts(importedDrafts)

      if (!cleaned.length) {
        setStatusMessage('Import file did not contain valid assignment rows.')
      } else {
        setDraftAssignments(cleaned)
        setDraftGroupNames(cleaned.map(() => ''))
        setStatusMessage(`Imported ${cleaned.length} assignment rows.`)
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Import failed.'
      setStatusMessage(message)
    } finally {
      event.target.value = ''
    }
  }

  const signIn = async () => {
    setIsBusy(true)
    setStatusMessage('Redirecting to Microsoft sign-in…')
    await instance.loginRedirect({ scopes: graphScopes })
  }

  const signOut = async () => {
    setIsBusy(true)
    setPolicies([])
    setSelectedPolicyIds(new Set())
    await instance.logoutRedirect()
  }

  const saveConfig = () => {
    const clientId = setupClientId.trim()
    const tenantId = setupTenantId.trim()
    if (!clientId || !tenantId) return
    onAuthConfigChange({ clientId, tenantId }, true)
  }

  return (
    <div className="app-shell">
      {deleteConfirmOpen && (
        <ConfirmDeleteModal
          policies={selectedPolicies}
          onConfirm={() => void runBulkDelete()}
          onCancel={() => setDeleteConfirmOpen(false)}
        />
      )}
      <header className="topbar">
        <div className="topbar-brand">
          <h1>Intune Overlord</h1>
          <p>Bulk import, export, add, and remove Intune assignments fast.</p>
        </div>
        <div className="auth-actions">
          {isAuthenticated ? (
            <>
              <span className="auth-user">{accounts[0]?.username}</span>
              <button type="button" className="secondary" onClick={() => void signOut()} disabled={isBusy}>
                Sign out
              </button>
            </>
          ) : clientIdConfigured ? (
            <>
              <button type="button" className="primary" onClick={() => void signIn()} disabled={isBusy}>
                Sign in
              </button>
              <button
                type="button"
                className="secondary"
                onClick={() => onAuthConfigChange({ clientId: '', tenantId: 'organizations' })}
                disabled={isBusy}
                title="Remove saved configuration and re-run setup"
              >
                Reset setup
              </button>
            </>
          ) : null}
        </div>
      </header>

      {!isAuthenticated && !clientIdConfigured && (
        <section className="setup-card">
          <h2 className="setup-card-title">First-time setup</h2>
          <p className="setup-card-intro">
            Register an app in Microsoft Entra, then paste the details below to connect.
          </p>
          <ol className="setup-steps">
            <li>Azure Portal → Entra ID → App registrations → <strong>New registration</strong></li>
            <li>
              Under <strong>Redirect URIs</strong>, add a <strong>Single-page application (SPA)</strong> URI:{' '}
              <code className="setup-uri">{window.location.origin}</code>
            </li>
            <li>
              Add delegated API permissions:{' '}
              <code>DeviceManagementConfiguration.ReadWrite.All</code>,{' '}
              <code>DeviceManagementManagedDevices.Read.All</code>,{' '}
              <code>DeviceManagementScripts.ReadWrite.All</code>,{' '}
              <code>Group.Read.All</code>
            </li>
            <li>Click <strong>Grant admin consent</strong> for your organisation</li>
          </ol>
          <div className="setup-fields">
            <input
              className="setup-field-input"
              value={setupTenantId}
              onChange={(e) => setSetupTenantId(e.target.value)}
              placeholder="Tenant ID or domain (e.g. contoso.onmicrosoft.com)"
            />
            <input
              className="setup-field-input"
              value={setupClientId}
              onChange={(e) => setSetupClientId(e.target.value)}
              placeholder="Client ID (Application ID)"
            />
            <button
              type="button"
              className="primary"
              onClick={saveConfig}
              disabled={!setupClientId.trim() || !setupTenantId.trim()}
            >
              Save &amp; Sign in
            </button>
          </div>
        </section>
      )}

      {!isAuthenticated && clientIdConfigured && (
        <section className="setup-hint setup-hint--ready">
          <span>
            App registered — client ID <code>{authConfig.clientId}</code>. Click <strong>Sign in</strong> to authenticate.
          </span>
        </section>
      )}

      <section className="controls-row">
        <button type="button" className="primary" onClick={runBulkAdd} disabled={isBusy || !isAuthenticated || !selectedPolicyIds.size}>
          Bulk add assignments
        </button>
        <button type="button" className="danger" onClick={runBulkRemove} disabled={isBusy || !isAuthenticated || !selectedPolicyIds.size}>
          Bulk remove assignments
        </button>
        <button
          type="button"
          className="danger danger--deep"
          onClick={() => setDeleteConfirmOpen(true)}
          disabled={isBusy || !isAuthenticated || !selectedPolicyIds.size}
          title="Permanently delete the selected policies from Intune"
        >
          ⚠️ Delete policies
        </button>
        <button
          type="button"
          className="danger danger--nuclear"
          onClick={handleDeleteAll}
          disabled={isBusy || !isAuthenticated || !policies.length}
          title="Select ALL loaded policies and permanently delete them from Intune"
        >
          ☢️ Delete ALL policies
        </button>
      </section>

      <main className="content-grid">
        <section className="panel">
          <div className="panel-header">
            <h2>Policies</h2>
            <div className="panel-header-actions">
              {selectedPolicyIds.size > 0 && (
                <span className="count-badge">{selectedPolicyIds.size} selected</span>
              )}
              {policies.length > 0 && (
                <button
                  type="button"
                  className="secondary panel-btn"
                  onClick={handleSelectAll}
                  disabled={isBusy}
                  title={allFilteredSelected ? 'Deselect all visible policies' : 'Select all visible policies'}
                >
                  {allFilteredSelected ? 'Deselect all' : 'Select all'}
                </button>
              )}
              <button
                type="button"
                className="secondary panel-btn"
                onClick={loadPolicies}
                disabled={isBusy || !isAuthenticated}
              >
                {isBusy ? 'Loading…' : policies.length > 0 ? 'Refresh' : 'Load policies'}
              </button>
            </div>
          </div>
          <input
            className="search"
            value={policySearch}
            onChange={(event) => setPolicySearch(event.target.value)}
            placeholder="Search policies"
          />
          <div className="policy-list">
            {filteredPolicies.length === 0 && <p className="empty">No policies found.</p>}
            {categoryOrder.map((kind) => {
              const group = groupedPolicies.get(kind) ?? []
              if (group.length === 0) return null
              const isCollapsed = collapsedCategories.has(kind)
              const selectedInGroup = group.filter((p) => selectedPolicyIds.has(p.id)).length
              const allSelected = selectedInGroup === group.length
              return (
                <div className="policy-category" key={kind}>
                  <div className="policy-category-header">
                    <button
                      type="button"
                      className="policy-category-toggle"
                      data-kind={kind}
                      onClick={() => toggleCategory(kind)}
                      aria-expanded={!isCollapsed}
                    >
                      <span className={`policy-category-chevron${isCollapsed ? '' : ' open'}`}>▶</span>
                      <span className="policy-category-name">{kind}</span>
                      <span className="policy-category-count">{group.length}</span>
                    </button>
                    <input
                      type="checkbox"
                      className="policy-category-check"
                      title={`Select all ${kind}`}
                      ref={(el) => { if (el) el.indeterminate = selectedInGroup > 0 && !allSelected }}
                      checked={allSelected}
                      onChange={() => toggleCategorySelection(kind)}
                    />
                  </div>
                  {!isCollapsed && (
                    <div className="policy-category-items">
                      {group.map((policy) => (
                        <label className={`policy-item${selectedPolicyIds.has(policy.id) ? ' selected' : ''}`} key={policy.id}>
                          <input
                            type="checkbox"
                            checked={selectedPolicyIds.has(policy.id)}
                            onChange={(event) => togglePolicy(policy.id, event.target.checked)}
                          />
                          <div>
                            <strong>{policy.displayName}</strong>
                          </div>
                        </label>
                      ))}
                    </div>
                  )}
                </div>
              )
            })}
          </div>
        </section>

        <section className="panel">
          <div className="panel-header">
            <div className="panel-tabs">
              <button
                type="button"
                className={`panel-tab${activeTab === 'assign' ? ' active' : ''}`}
                onClick={() => setActiveTab('assign')}
              >
                Assign
              </button>
              <button
                type="button"
                className={`panel-tab${activeTab === 'view' ? ' active' : ''}`}
                onClick={switchToViewTab}
              >
                Current
              </button>
            </div>
            <div className="panel-header-actions">
              {activeTab === 'assign' ? (
                <>
                  <button
                    type="button"
                    className="secondary panel-btn"
                    onClick={exportAssignments}
                    disabled={isBusy || !isAuthenticated}
                    title="Export current assignments from selected policies"
                  >
                    Export
                  </button>
                  <label className="secondary file-input panel-btn" title="Import assignments from CSV or JSON">
                    Import
                    <input type="file" accept=".csv,.json" onChange={importAssignments} disabled={isBusy} />
                  </label>
                  <button type="button" className="secondary panel-btn" onClick={appendDraft}>
                    + Row
                  </button>
                </>
              ) : (
                <button
                  type="button"
                  className="secondary panel-btn"
                  onClick={() => void loadExistingAssignments()}
                  disabled={isViewLoading || !isAuthenticated || !selectedPolicyIds.size}
                >
                  {isViewLoading ? 'Loading…' : 'Refresh'}
                </button>
              )}
            </div>
          </div>

          {activeTab === 'assign' ? (
            <>
              <div className="draft-list">
                {draftAssignments.map((draft, index) => (
                  <div className="draft-row" key={`${index}-${draft.groupId}`}>
                    <GroupPicker
                      resolvedName={draftGroupNames[index] ?? ''}
                      groupId={draft.groupId}
                      disabled={isBusy}
                      getToken={getToken}
                      onSelect={(id, name) => {
                        updateDraft(index, { ...draft, groupId: id })
                        setDraftGroupNames((current) => current.map((n, i) => (i === index ? name : n)))
                      }}
                    />
                    <select
                      value={draft.mode}
                      onChange={(event) =>
                        updateDraft(index, {
                          ...draft,
                          mode: normalizeMode(event.target.value),
                        })
                      }
                    >
                      <option value="include">Include</option>
                      <option value="exclude">Exclude</option>
                    </select>
                    <button type="button" className="danger" onClick={() => removeDraft(index)}>
                      Remove
                    </button>
                  </div>
                ))}
              </div>
              <p className="helper">Import supports CSV or JSON with columns/keys: groupId, mode(include|exclude)</p>
            </>
          ) : (
            <div className="existing-list">
              {!selectedPolicyIds.size && (
                <p className="empty">Select policies on the left to view their assignments.</p>
              )}
              {selectedPolicyIds.size > 0 && isViewLoading && (
                <p className="empty">Loading…</p>
              )}
              {selectedPolicyIds.size > 0 && !isViewLoading && existingAssignments.map((block) => (
                <div className="existing-block" key={block.policyId}>
                  <div className="existing-block-header">
                    <span className="existing-policy-name">{block.policyName}</span>
                    <span className="existing-policy-kind">{block.policyKind}</span>
                  </div>
                  {block.assignments.length === 0 ? (
                    <p className="existing-none">No assignments</p>
                  ) : (
                    <ul className="existing-assignments">
                      {block.assignments.map((a) => (
                        <li key={`${a.groupId}-${a.mode}`} className="existing-assignment-row">
                          <span className={`existing-mode-badge existing-mode-badge--${a.mode}`}>
                            {a.mode}
                          </span>
                          <span className="existing-group-name">
                            {a.groupName || <span className="existing-group-id">{a.groupId}</span>}
                          </span>
                          {a.groupName && (
                            <span className="existing-group-id">{a.groupId}</span>
                          )}
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              ))}
            </div>
          )}
        </section>
      </main>

      <footer className="statusbar">
        <span>{isBusy ? 'Working...' : 'Ready'}</span>
        <span>{statusMessage}</span>
      </footer>
    </div>
  )
}

export default App
