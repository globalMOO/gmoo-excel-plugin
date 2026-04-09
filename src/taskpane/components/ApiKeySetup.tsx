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
  Link,
} from "@fluentui/react-components";
import {
  PlugConnected20Regular,
  ChevronDown20Regular,
  ChevronRight20Regular,
} from "@fluentui/react-icons";
import { GmooClient, GmooApiError } from "../services/gmooApi";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  inputRow: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  inputField: {
    flexGrow: 1,
  },
  settingsToggle: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    cursor: "pointer",
    color: tokens.colorNeutralForeground3,
    ":hover": {
      color: tokens.colorNeutralForeground1,
    },
  },
  settingsPanel: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    paddingLeft: "8px",
    borderLeft: `2px solid ${tokens.colorNeutralStroke2}`,
  },
});

interface ApiKeySetupProps {
  apiKey: string;
  apiUrl: string;
  defaultApiUrl: string;
  onApiKeyChange: (key: string) => Promise<void>;
  onApiUrlChange: (url: string) => Promise<void>;
  onNext: () => void;
}

export const ApiKeySetup: React.FC<ApiKeySetupProps> = ({
  apiKey,
  apiUrl,
  defaultApiUrl,
  onApiKeyChange,
  onApiUrlChange,
  onNext,
}) => {
  const styles = useStyles();
  const [inputValue, setInputValue] = useState(apiKey);
  const [urlValue, setUrlValue] = useState(apiUrl);
  const [showSettings, setShowSettings] = useState(apiUrl !== defaultApiUrl);
  const [isValidating, setIsValidating] = useState(false);
  const [isConnected, setIsConnected] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleConnect = async () => {
    const key = inputValue.trim();
    if (!key) {
      setError("Please enter an API key.");
      return;
    }

    setIsValidating(true);
    setError(null);
    setIsConnected(false);

    try {
      await onApiUrlChange(urlValue);
      const resolvedUrl = resolveUrlForValidation(urlValue, defaultApiUrl);
      const tempClient = new GmooClient(key, resolvedUrl);
      await tempClient.getModels();
      await onApiKeyChange(key);
      setIsConnected(true);
    } catch (err) {
      if (err instanceof GmooApiError && err.status === 401) {
        setError("Invalid API key. Please check your key and try again.");
      } else if (err instanceof GmooApiError) {
        setError(`API error (${err.status}): ${err.apiError?.message ?? "Unknown error"}`);
      } else {
        setError(`Connection failed: ${err instanceof Error ? err.message : "Unknown error"}`);
      }
    } finally {
      setIsValidating(false);
    }
  };

  // Auto-validate if we already have a saved key
  React.useEffect(() => {
    if (apiKey && !isConnected && !isValidating) {
      setInputValue(apiKey);
      handleConnect();
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div className={styles.container}>
      <Card>
        <CardHeader
          image={<PlugConnected20Regular />}
          header={<Text weight="semibold">Connect to globalMOO</Text>}
          description="Enter your API key to connect to the globalMOO service."
        />
      </Card>

      <Text size={200}>
        Get your API key from your globalMOO account at app.globalmoo.com
      </Text>

      <div className={styles.inputRow}>
        <Input
          className={styles.inputField}
          type="password"
          placeholder="Enter API key..."
          value={inputValue}
          onChange={(_, data) => {
            setInputValue(data.value);
            setIsConnected(false);
            setError(null);
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") handleConnect();
          }}
          disabled={isValidating}
        />
        <Button
          appearance="primary"
          onClick={handleConnect}
          disabled={isValidating || !inputValue.trim()}
        >
          {isValidating ? <Spinner size="tiny" /> : "Connect"}
        </Button>
      </div>

      <div
        className={styles.settingsToggle}
        onClick={() => setShowSettings(!showSettings)}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => { if (e.key === "Enter") setShowSettings(!showSettings); }}
      >
        {showSettings ? <ChevronDown20Regular /> : <ChevronRight20Regular />}
        <Text size={200}>API Settings</Text>
      </div>

      {showSettings && (
        <div className={styles.settingsPanel}>
          <Text size={200}>API URL</Text>
          <Input
            size="small"
            value={urlValue}
            onChange={(_, data) => {
              setUrlValue(data.value);
              setIsConnected(false);
              setError(null);
            }}
            placeholder={defaultApiUrl}
          />
          {urlValue !== defaultApiUrl && (
            <Link
              as="button"
              onClick={() => {
                setUrlValue(defaultApiUrl);
                setIsConnected(false);
              }}
            >
              Reset to default
            </Link>
          )}
        </div>
      )}

      {isConnected && (
        <MessageBar intent="success">
          <MessageBarBody>
            <MessageBarTitle>Connected</MessageBarTitle>
            Successfully connected to globalMOO.
          </MessageBarBody>
        </MessageBar>
      )}

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Connection Failed</MessageBarTitle>
            {error}
          </MessageBarBody>
        </MessageBar>
      )}

      <Button
        appearance="primary"
        onClick={onNext}
        disabled={!isConnected}
        style={{ alignSelf: "flex-end" }}
      >
        Next
      </Button>
    </div>
  );
};

// Duplicates the resolution logic from useGmooClient for validation
function resolveUrlForValidation(displayUrl: string, defaultUrl: string): string {
  const normalized = displayUrl.trim().replace(/\/+$/, "") + "/";
  const defaultNormalized = defaultUrl.trim().replace(/\/+$/, "") + "/";
  if (normalized === defaultNormalized) {
    return "https://localhost:3001/api/";
  }
  return normalized;
}
