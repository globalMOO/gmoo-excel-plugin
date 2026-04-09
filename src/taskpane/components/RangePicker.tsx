import React, { useState, useCallback } from "react";
import {
  makeStyles,
  Button,
  Input,
  Text,
} from "@fluentui/react-components";
import { CursorClick20Regular } from "@fluentui/react-icons";
import { readSelectedRange } from "../services/excelService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  addressInput: {
    flexGrow: 1,
  },
});

interface RangePickerProps {
  label: string;
  value: string;
  onChange: (address: string) => void;
}

export const RangePicker: React.FC<RangePickerProps> = ({
  label,
  value,
  onChange,
}) => {
  const styles = useStyles();
  const [isPicking, setIsPicking] = useState(false);

  const handlePick = useCallback(async () => {
    setIsPicking(true);
    try {
      const address = await readSelectedRange();
      onChange(address);
    } catch {
      // User cancelled or error
    } finally {
      setIsPicking(false);
    }
  }, [onChange]);

  return (
    <div>
      <Text size={200} weight="semibold">
        {label}
      </Text>
      <div className={styles.container}>
        <Input
          className={styles.addressInput}
          size="small"
          value={value}
          onChange={(_, data) => onChange(data.value)}
          placeholder="e.g., Sheet1!B2:B5"
        />
        <Button
          icon={<CursorClick20Regular />}
          size="small"
          appearance="subtle"
          onClick={handlePick}
          disabled={isPicking}
          title="Select range in Excel"
        >
          {isPicking ? "Click cells..." : "Pick"}
        </Button>
      </div>
    </div>
  );
};
