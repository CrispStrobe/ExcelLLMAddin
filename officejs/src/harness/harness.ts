// Dev harness entry: install the Office mock, then load the real task pane.
// Import order matters — the mock defines globalThis.Office/OfficeRuntime before
// taskpane.ts runs Office.onReady at module load.
import "./officeMock";
import "../taskpane/taskpane";
