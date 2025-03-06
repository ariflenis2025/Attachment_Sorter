import React from "react";
import { makeStyles, shorthands } from "@fluentui/react-components";

interface BarProps {
  percent: any; // The percentage value for the progress bar (0-100)
}

const useStyles = makeStyles({
  progressBarContainer: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    width: "150px", // Circular bar width/height
    height: "150px",
    borderRadius: "50%",
    background:'#5bb3ff', // Subtle inner circle
    position: "relative",
    boxShadow: "0 4px 8px rgba(0, 0, 0, 0.1)", // Soft shadow for a modern look
  },
  progressBar: {
    position: "absolute",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    borderRadius: "50%",
    background: `conic-gradient(#4caf50 calc(var(--progress-value) * 1%), #e0e0e0 0)`, // Green progress + gray remainder
    transform: "rotate(-90deg)", // Rotate to start from the top
    clipPath: "inset(0)", // Ensures clipping works consistently
    transition: "background 0.5s ease-in-out", // Smooth animation for changes
  },
  progressText: {
    position: "absolute",
    fontSize: "1.2rem",
    fontWeight: 600,
    color: "#333",
    textAlign: "center",
  },
});
const Bar: React.FC<BarProps> = ({ percent }) => {
  const styles = useStyles();

  return (
    <div
      className={styles.progressBarContainer}
      style={{ "--progress-value": `${percent}%` } as React.CSSProperties} // Dynamic percentage
    >
      <div className={styles.progressBar}></div>
      <div className={styles.progressText}>{percent}</div>
    </div>
  );
};

export default Bar;
