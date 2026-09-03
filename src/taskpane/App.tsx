import * as React from "react";
import {
  Body1,
  Caption1,
  Radio,
  RadioGroup,
  Subtitle1,
  Title3,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import { loadOrientation, persistOrientation, type Orientation } from "../shared/orientation";

/* global Excel */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "20px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  code: {
    fontFamily: tokens.fontFamilyMonospace,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "2px 6px",
    borderRadius: tokens.borderRadiusSmall,
  },
});

/** Forces a full recalculation so open FACTORIALROW cells re-spill in the new orientation. */
async function recalculateWorkbook(): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.application.calculate(Excel.CalculationType.full);
    await context.sync();
  });
}

export const App: React.FC = () => {
  const styles = useStyles();
  const [orientation, setOrientationState] = React.useState<Orientation>("row");

  React.useEffect(() => {
    void loadOrientation().then(setOrientationState);
  }, []);

  const handleChange = React.useCallback(async (next: Orientation) => {
    setOrientationState(next);
    try {
      await persistOrientation(next);
      await recalculateWorkbook();
    } catch (error) {
      // Persisting/recalculating is best-effort; keep the UI responsive on failure.
      console.error("Failed to apply orientation change:", error);
    }
  }, []);

  return (
    <main className={styles.root}>
      <Title3>Velixo FACTORIALROW</Title3>

      <div className={styles.section}>
        <Body1>
          Use <span className={styles.code}>=TESTVELIXO.FACTORIALROW(N)</span> in a cell to spill{" "}
          <span className={styles.code}>[0!, 1!, 2!, ..., N!]</span>.
        </Body1>
        <Caption1>Supports N from 0 to 500 without losing precision.</Caption1>
      </div>

      <div className={styles.section}>
        <Subtitle1>Output orientation</Subtitle1>
        <RadioGroup
          value={orientation}
          onChange={(_, data) => void handleChange(data.value as Orientation)}
        >
          <Radio value="row" label="Row (spills right)" />
          <Radio value="column" label="Column (spills down)" />
        </RadioGroup>
        <Caption1>Changing this recalculates the workbook so results re-orient immediately.</Caption1>
      </div>
    </main>
  );
};
