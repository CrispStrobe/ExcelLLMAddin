// Minimal ambient declaration for the OfficeRuntime.storage API.
// @types/office-js (1.0.x) does not declare the OfficeRuntime global, and
// @types/custom-functions-runtime declares only CustomFunctions. This covers the
// shared async storage we use from both the task pane and the functions runtime.
declare namespace OfficeRuntime {
  const storage: {
    getItem(key: string): Promise<string | null>;
    setItem(key: string, value: string): Promise<void>;
    removeItem(key: string): Promise<void>;
    getItems(keys: string[]): Promise<{ [key: string]: string | null }>;
    setItems(items: { [key: string]: string }): Promise<void>;
    removeItems(keys: string[]): Promise<void>;
  };
}
