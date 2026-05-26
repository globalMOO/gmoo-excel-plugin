import React, { useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Input,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Card,
  CardHeader,
  Field,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Badge,
  Tooltip,
} from "@fluentui/react-components";
import {
  PlugConnected20Regular,
  Add20Regular,
  Edit20Regular,
  Delete20Regular,
  CheckmarkCircle20Filled,
  Copy20Regular,
  ArrowClockwise20Regular,
} from "@fluentui/react-icons";
import { GmooClient } from "../services/gmooApi";
import { diagnoseConnection, buildCertTrustCommand } from "../services/connectionDiagnostics";
import type { Connection, NewConnectionInput } from "../types/connection";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  list: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  row: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  rowActive: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
  },
  rowMain: {
    flexGrow: 1,
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    minWidth: 0, // allow truncation
  },
  rowLabel: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  rowUrl: {
    color: tokens.colorNeutralForeground3,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  rowActions: {
    display: "flex",
    gap: "4px",
  },
  formFields: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  emptyState: {
    padding: "20px",
    border: `1px dashed ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
  },
  certBlock: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "8px",
  },
  certCommand: {
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "11px",
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "8px",
    borderRadius: tokens.borderRadiusMedium,
    overflowX: "auto",
    whiteSpace: "pre-wrap",
    wordBreak: "break-all",
    margin: 0,
  },
  certActions: {
    display: "flex",
    gap: "8px",
  },
});

export interface ConnectionSetupProps {
  connections: Connection[];
  activeConnection: Connection | null;
  onSetActive: (id: string) => Promise<void>;
  onAdd: (input: NewConnectionInput) => Promise<Connection>;
  onUpdate: (id: string, patch: Partial<Connection>) => Promise<Connection>;
  onDelete: (id: string) => Promise<void>;
  onNext: () => void;
  /** Optional banner shown above the list (used to surface activation errors). */
  banner?: { intent: "success" | "error" | "info"; title: string; body?: string } | null;
  onDismissBanner?: () => void;
}

export const ConnectionSetup: React.FC<ConnectionSetupProps> = ({
  connections,
  activeConnection,
  onSetActive,
  onAdd,
  onUpdate,
  onDelete,
  onNext,
  banner,
  onDismissBanner,
}) => {
  const styles = useStyles();
  const [editTarget, setEditTarget] = useState<Connection | "new" | null>(null);
  const [validateState, setValidateState] = useState<{
    /** Connection id currently under validation, or null if no test is running. */
    inFlightId: string | null;
    error: string | null;
    failureKind: "api" | "network" | null;
    /**
     * Result of follow-up no-cors probe when failureKind is "network".
     * true  = server is reachable + TLS-trusted, so the original failure was CORS.
     * false = probe also failed (cert untrusted / DNS / network unreachable).
     * null  = probe didn't run (dev mode, or failure wasn't network-class).
     */
    corsReachable: boolean | null;
    lastTested: Connection | null;
    validatedId: string | null;
  }>({
    inFlightId: null,
    error: null,
    failureKind: null,
    corsReachable: null,
    lastTested: null,
    validatedId: null,
  });
  const [copied, setCopied] = useState(false);

  const validate = async (conn: Connection) => {
    setValidateState({
      inFlightId: conn.id,
      error: null,
      failureKind: null,
      corsReachable: null,
      lastTested: conn,
      validatedId: null,
    });
    setCopied(false);

    const host = conn.apiUrl.replace(/\/+$/, "").replace(/\/api$/i, "");
    // Mirror useGmooClient's dev rewrite so the Connect-button validation
    // also goes through the webpack dev proxy and avoids the API's CORS
    // whitelist when served from https://localhost:3000.
    const isDev =
      typeof window !== "undefined" &&
      window.location.hostname === "localhost" &&
      host === "https://app.globalmoo.com";
    const base = isDev
      ? `${window.location.protocol}//${window.location.host}/api/`
      : host + "/api/";

    const client = new GmooClient(conn.apiKey, base);
    const result = await diagnoseConnection({
      client,
      probeHost: isDev ? null : host,
    });

    // Drop the result if the user kicked off a different test in the
    // meantime — otherwise we'd clobber the newer in-flight spinner.
    setValidateState((prev) =>
      prev.inFlightId !== conn.id
        ? prev
        : result.ok
        ? {
            inFlightId: null,
            error: null,
            failureKind: null,
            corsReachable: null,
            lastTested: conn,
            validatedId: conn.id,
          }
        : {
            inFlightId: null,
            error: result.error,
            failureKind: result.failureKind,
            corsReachable: result.corsReachable,
            lastTested: conn,
            validatedId: null,
          }
    );
  };

  const copyCertTrustCommand = async () => {
    if (!validateState.lastTested) return;
    const cmd = buildCertTrustCommand(validateState.lastTested.apiUrl);
    try {
      await navigator.clipboard.writeText(cmd);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    } catch {
      // Clipboard API can fail in restricted contexts — fall back to a manual
      // textarea select. Silent failure is fine; the command is visible.
    }
  };

  const retryValidate = () => {
    if (validateState.lastTested) {
      void validate(validateState.lastTested);
    }
  };

  const isEmpty = connections.length === 0;
  const canProceed = !!activeConnection && !!activeConnection.apiKey;

  return (
    <div className={styles.container}>
      <Card>
        <CardHeader
          image={<PlugConnected20Regular />}
          header={<Text weight="semibold">Connections</Text>}
          description="Manage globalMOO API connections. Pick one to use in this workbook."
        />
      </Card>

      {banner && (
        <MessageBar intent={banner.intent} onClick={onDismissBanner}>
          <MessageBarBody>
            <MessageBarTitle>{banner.title}</MessageBarTitle>
            {banner.body}
          </MessageBarBody>
        </MessageBar>
      )}

      {isEmpty ? (
        <div className={styles.emptyState}>
          <Text>No connections yet.</Text>
          <Text size={200} className={styles.hint}>
            Add your first connection below. If you signed up at globalmoo.com,
            you can also click the activation link in your welcome email to
            skip this step.
          </Text>
          <Button
            appearance="primary"
            icon={<Add20Regular />}
            onClick={() => setEditTarget("new")}
          >
            Add Connection
          </Button>
        </div>
      ) : (
        <>
          <div className={styles.list}>
            {connections.map((c) => (
              <ConnectionRow
                key={c.id}
                connection={c}
                isActive={c.id === activeConnection?.id}
                isValidating={validateState.inFlightId === c.id}
                onSetActive={() => onSetActive(c.id)}
                onEdit={() => setEditTarget(c)}
                onDelete={() => onDelete(c.id)}
                onTest={() => validate(c)}
              />
            ))}
          </div>

          <Button
            icon={<Add20Regular />}
            onClick={() => setEditTarget("new")}
            style={{ alignSelf: "flex-start" }}
          >
            Add Connection
          </Button>
        </>
      )}

      {validateState.error && validateState.failureKind === "api" && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Test failed</MessageBarTitle>
            {validateState.error}
          </MessageBarBody>
        </MessageBar>
      )}

      {validateState.error &&
        validateState.failureKind === "network" &&
        validateState.corsReachable === true &&
        validateState.lastTested && (() => {
          // The origin the customer's server needs to allow is whatever this
          // task pane is served from — typically https://globalmoo.github.io,
          // but anyone hosting a fork (or running dev) sees the right value.
          const origin =
            typeof window !== "undefined" ? window.location.origin : "https://globalmoo.github.io";
          return (
          <MessageBar intent="warning">
            <MessageBarBody>
              <MessageBarTitle>
                Server reachable, but the browser blocked the response (CORS)
              </MessageBarTitle>
              <div className={styles.certBlock}>
                <Text size={200}>
                  Your API server responded, but didn't allow this add-in's
                  origin (<code>{origin}</code>). Add these headers to your API
                  server's responses:
                </Text>
                <pre className={styles.certCommand}>
{`Access-Control-Allow-Origin: ${origin}
Access-Control-Allow-Headers: Authorization, Content-Type
Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS`}
                </pre>
                <details>
                  <summary>nginx snippet</summary>
                  <pre className={styles.certCommand}>
{`add_header Access-Control-Allow-Origin  "${origin}" always;
add_header Access-Control-Allow-Headers "Authorization, Content-Type" always;
add_header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS" always;
if ($request_method = OPTIONS) { return 204; }`}
                  </pre>
                </details>
                <details>
                  <summary>Apache snippet</summary>
                  <pre className={styles.certCommand}>
{`Header always set Access-Control-Allow-Origin  "${origin}"
Header always set Access-Control-Allow-Headers "Authorization, Content-Type"
Header always set Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS"
RewriteEngine On
RewriteCond %{REQUEST_METHOD} OPTIONS
RewriteRule ^ - [R=204,L]`}
                  </pre>
                </details>
                <details>
                  <summary>IIS web.config snippet</summary>
                  <pre className={styles.certCommand}>
{`<system.webServer>
  <httpProtocol>
    <customHeaders>
      <add name="Access-Control-Allow-Origin"  value="${origin}" />
      <add name="Access-Control-Allow-Headers" value="Authorization, Content-Type" />
      <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, OPTIONS" />
    </customHeaders>
  </httpProtocol>
</system.webServer>`}
                  </pre>
                </details>
                <div className={styles.certActions}>
                  <Button
                    size="small"
                    icon={<ArrowClockwise20Regular />}
                    appearance="primary"
                    onClick={retryValidate}
                  >
                    Retry
                  </Button>
                </div>
                <details>
                  <summary>Technical details</summary>
                  <Text size={100} className={styles.hint}>
                    {validateState.error}
                  </Text>
                </details>
              </div>
            </MessageBarBody>
          </MessageBar>
          );
        })()}

      {validateState.error &&
        validateState.failureKind === "network" &&
        validateState.corsReachable !== true &&
        validateState.lastTested && (
          <MessageBar intent="warning">
            <MessageBarBody>
              <MessageBarTitle>
                Couldn't reach{" "}
                {(() => {
                  try {
                    return new URL(validateState.lastTested.apiUrl).hostname;
                  } catch {
                    return validateState.lastTested.apiUrl;
                  }
                })()}
              </MessageBarTitle>
              <div className={styles.certBlock}>
                <Text size={200}>
                  {validateState.corsReachable === false
                    ? "The server didn't respond at all. The most common cause is a self-signed TLS certificate that Windows doesn't trust. Open PowerShell on this machine and run:"
                    : "This usually means the server's TLS certificate isn't trusted by Windows (also fires on DNS / network failures). Open PowerShell on this machine and run:"}
                </Text>
                <pre className={styles.certCommand}>
                  {buildCertTrustCommand(validateState.lastTested.apiUrl)}
                </pre>
                <Text size={200} className={styles.hint}>
                  The installer reports back whether it imported a cert, found
                  the host unreachable, or couldn't resolve DNS.
                </Text>
                <div className={styles.certActions}>
                  <Button
                    size="small"
                    icon={<Copy20Regular />}
                    onClick={copyCertTrustCommand}
                  >
                    {copied ? "Copied!" : "Copy command"}
                  </Button>
                  <Button
                    size="small"
                    icon={<ArrowClockwise20Regular />}
                    appearance="primary"
                    onClick={retryValidate}
                  >
                    Retry
                  </Button>
                </div>
                <details>
                  <summary>Technical details</summary>
                  <Text size={100} className={styles.hint}>
                    {validateState.error}
                  </Text>
                </details>
              </div>
            </MessageBarBody>
          </MessageBar>
        )}

      {validateState.validatedId && (
        <MessageBar intent="success">
          <MessageBarBody>
            <MessageBarTitle>Connection works</MessageBarTitle>
            Successfully reached the globalMOO API.
          </MessageBarBody>
        </MessageBar>
      )}

      <Button
        appearance="primary"
        onClick={onNext}
        disabled={!canProceed}
        style={{ alignSelf: "flex-end" }}
      >
        Next
      </Button>

      {editTarget && (
        <ConnectionDialog
          target={editTarget}
          existingLabels={new Set(
            connections
              .filter((c) => editTarget === "new" || c.id !== editTarget.id)
              .map((c) => c.label)
          )}
          onSave={async (input) => {
            if (editTarget === "new") {
              const created = await onAdd(input);
              await onSetActive(created.id);
            } else {
              await onUpdate(editTarget.id, input);
            }
            setEditTarget(null);
          }}
          onCancel={() => setEditTarget(null)}
        />
      )}
    </div>
  );
};

// --- Sub-components -------------------------------------------------------

interface ConnectionRowProps {
  connection: Connection;
  isActive: boolean;
  isValidating: boolean;
  onSetActive: () => void;
  onEdit: () => void;
  onDelete: () => void;
  onTest: () => void;
}

const ConnectionRow: React.FC<ConnectionRowProps> = ({
  connection,
  isActive,
  isValidating,
  onSetActive,
  onEdit,
  onDelete,
  onTest,
}) => {
  const styles = useStyles();
  const [confirmingDelete, setConfirmingDelete] = useState(false);

  const hostname = (() => {
    try {
      return new URL(connection.apiUrl).hostname;
    } catch {
      return connection.apiUrl;
    }
  })();

  return (
    <div className={isActive ? `${styles.row} ${styles.rowActive}` : styles.row}>
      <div className={styles.rowMain}>
        <div className={styles.rowLabel}>
          {isActive && <CheckmarkCircle20Filled style={{ color: "var(--colorBrandForeground1)" }} />}
          <Text weight="semibold">{connection.label}</Text>
          {!connection.apiKey && (
            <Badge appearance="outline" size="small">
              No key
            </Badge>
          )}
          {connection.source === "activation" && (
            <Badge appearance="tint" size="small">
              Activated
            </Badge>
          )}
        </div>
        <Text size={200} className={styles.rowUrl}>
          {hostname}
        </Text>
      </div>
      <div className={styles.rowActions}>
        {!isActive && (
          <Button size="small" onClick={onSetActive}>
            Set Active
          </Button>
        )}
        <Tooltip content="Test connection" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            disabled={!connection.apiKey || isValidating}
            onClick={onTest}
          >
            {isValidating ? <Spinner size="tiny" /> : "Test"}
          </Button>
        </Tooltip>
        <Tooltip content="Edit" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Edit20Regular />}
            onClick={onEdit}
          />
        </Tooltip>
        <Tooltip content="Delete" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Delete20Regular />}
            onClick={() => setConfirmingDelete(true)}
          />
        </Tooltip>
      </div>

      <Dialog open={confirmingDelete} onOpenChange={(_, data) => setConfirmingDelete(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Delete connection?</DialogTitle>
            <DialogContent>
              <Text>
                Remove "{connection.label}" from this device? The API key will
                be cleared from local storage but is not revoked on the server.
              </Text>
            </DialogContent>
            <DialogActions>
              <Button onClick={() => setConfirmingDelete(false)}>Cancel</Button>
              <Button
                appearance="primary"
                onClick={() => {
                  setConfirmingDelete(false);
                  onDelete();
                }}
              >
                Delete
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

interface ConnectionDialogProps {
  target: Connection | "new";
  existingLabels: Set<string>;
  onSave: (input: NewConnectionInput) => Promise<void>;
  onCancel: () => void;
}

const ConnectionDialog: React.FC<ConnectionDialogProps> = ({
  target,
  existingLabels,
  onSave,
  onCancel,
}) => {
  const styles = useStyles();
  const isNew = target === "new";
  const initial = isNew
    ? { label: "", apiUrl: "https://app.globalmoo.com", apiKey: "" }
    : { label: target.label, apiUrl: target.apiUrl, apiKey: target.apiKey };

  const [label, setLabel] = useState(initial.label);
  const [apiUrl, setApiUrl] = useState(initial.apiUrl);
  const [apiKey, setApiKey] = useState(initial.apiKey);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const labelError = (() => {
    if (!label.trim()) return "Label is required.";
    if (existingLabels.has(label.trim())) return "Another connection already uses this label.";
    return null;
  })();

  const urlError = (() => {
    const trimmed = apiUrl.trim();
    if (!trimmed) return "API URL is required.";
    try {
      const u = new URL(trimmed);
      if (u.protocol !== "https:") return "URL must start with https://";
    } catch {
      return "Not a valid URL.";
    }
    return null;
  })();

  const canSave = !labelError && !urlError && !saving;

  const handleSave = async () => {
    if (!canSave) return;
    setSaving(true);
    setError(null);
    try {
      await onSave({
        label: label.trim(),
        apiUrl: apiUrl.trim(),
        apiKey,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Save failed.");
      setSaving(false);
    }
  };

  return (
    <Dialog open={true} onOpenChange={(_, data) => !data.open && onCancel()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{isNew ? "Add Connection" : "Edit Connection"}</DialogTitle>
          <DialogContent>
            <div className={styles.formFields}>
              <Field
                label="Label"
                validationState={labelError ? "error" : "none"}
                validationMessage={labelError ?? undefined}
              >
                <Input
                  value={label}
                  onChange={(_, data) => setLabel(data.value)}
                  placeholder="e.g. globalMOO Cloud"
                />
              </Field>
              <Field
                label="API URL"
                validationState={urlError ? "error" : "none"}
                validationMessage={urlError ?? undefined}
                hint="The base URL for your globalMOO server (no /api/ suffix)."
              >
                <Input
                  value={apiUrl}
                  onChange={(_, data) => setApiUrl(data.value)}
                  placeholder="https://app.globalmoo.com"
                />
              </Field>
              <Field
                label="API Key"
                hint="Leave blank to save a key-less placeholder. You can fill it in later."
              >
                <Input
                  type="password"
                  value={apiKey}
                  onChange={(_, data) => setApiKey(data.value)}
                  placeholder="gm_live_..."
                />
              </Field>
              {error && (
                <MessageBar intent="error">
                  <MessageBarBody>{error}</MessageBarBody>
                </MessageBar>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button onClick={onCancel} disabled={saving}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={handleSave} disabled={!canSave}>
              {saving ? <Spinner size="tiny" /> : "Save"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
