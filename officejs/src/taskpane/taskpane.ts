import "./taskpane.css";
import { PROVIDERS, getProvider } from "../core/providers";
import { loadSettings, saveSettings } from "../core/config";
import { runPrompt, listModels, LlmSettings } from "../core/llm";
import { runAgent } from "../core/agent";
import { EXCEL_TOOLS, executeExcelTool } from "../excelTools";
import { browserFetch as fetchLike } from "../browserFetch";

/* global Office, document, window */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    void init();
  }
});

let allModels: string[] = [];

async function init(): Promise<void> {
  populateProviders();
  setForm(await loadSettings());
  byId<HTMLButtonElement>("save").onclick = onSave;
  byId<HTMLButtonElement>("test").onclick = onTest;
  byId<HTMLButtonElement>("refreshModels").onclick = onRefreshModels;
  byId<HTMLSelectElement>("provider").onchange = onProviderChange;
  byId<HTMLInputElement>("modelFilter").oninput = () =>
    renderModels(byId<HTMLInputElement>("modelFilter").value);
  byId<HTMLSelectElement>("modelSelect").onchange = onPickModel;
  byId<HTMLButtonElement>("reload").onclick = () => window.location.reload();
  byId<HTMLButtonElement>("agentRun").onclick = onAgentRun;
  updateKeyHint();
}

async function onAgentRun(): Promise<void> {
  const input = byId<HTMLTextAreaElement>("agentInput").value.trim();
  const log = byId<HTMLDivElement>("agentLog");
  if (!input) {
    log.textContent = "Type an instruction first.";
    return;
  }
  log.textContent = "Working…";
  try {
    const res = await runAgent(input, EXCEL_TOOLS, readForm(), { fetch: fetchLike }, executeExcelTool);
    const lines = res.steps.map((s) => `• ${s.tool}(${clip(JSON.stringify(s.args))}) → ${clip(s.result)}`);
    lines.push("", res.finalText || "(done)");
    log.textContent = lines.join("\n");
  } catch (e) {
    log.textContent = errText(e);
  }
}

function clip(s: string): string {
  return s.length > 90 ? s.slice(0, 90) + "…" : s;
}

function populateProviders(): void {
  const sel = byId<HTMLSelectElement>("provider");
  sel.innerHTML = "";
  for (const spec of Object.values(PROVIDERS)) {
    const opt = document.createElement("option");
    opt.value = spec.id;
    opt.textContent = spec.label;
    sel.appendChild(opt);
  }
}

function setForm(s: LlmSettings): void {
  byId<HTMLSelectElement>("provider").value = s.provider;
  byId<HTMLInputElement>("model").value = s.model || "";
  byId<HTMLInputElement>("apiKey").value = s.apiKey || "";
  byId<HTMLInputElement>("baseUrl").value = s.baseUrl || "";
  byId<HTMLInputElement>("proxyUrl").value = s.proxyUrl || "";
  byId<HTMLInputElement>("embedModel").value = s.embedModel || "";
  byId<HTMLTextAreaElement>("systemPrompt").value = s.systemPrompt || "";
}

function readForm(): LlmSettings {
  return {
    provider: byId<HTMLSelectElement>("provider").value,
    model: byId<HTMLInputElement>("model").value.trim(),
    apiKey: byId<HTMLInputElement>("apiKey").value.trim(),
    baseUrl: byId<HTMLInputElement>("baseUrl").value.trim(),
    proxyUrl: byId<HTMLInputElement>("proxyUrl").value.trim(),
    embedModel: byId<HTMLInputElement>("embedModel").value.trim(),
    systemPrompt: byId<HTMLTextAreaElement>("systemPrompt").value.trim(),
  };
}

async function onSave(): Promise<void> {
  try {
    await saveSettings(readForm());
    setStatus("Saved.", "ok");
  } catch (e) {
    setStatus(errText(e), "err");
  }
}

async function onTest(): Promise<void> {
  setStatus("Testing…", "");
  try {
    const reply = await runPrompt("Reply with exactly: Hello from Excel", readForm(), { fetch: fetchLike });
    setStatus("OK: " + reply, "ok");
  } catch (e) {
    setStatus(errText(e), "err");
  }
}

async function onRefreshModels(): Promise<void> {
  setStatus("Loading models…", "");
  try {
    allModels = await listModels(readForm(), { fetch: fetchLike });
    byId<HTMLInputElement>("modelFilter").classList.remove("hidden");
    byId<HTMLSelectElement>("modelSelect").classList.remove("hidden");
    renderModels(byId<HTMLInputElement>("modelFilter").value);
    setStatus(`${allModels.length} models — filter, then click one to select.`, "ok");
  } catch (e) {
    setStatus(errText(e), "err");
  }
}

function renderModels(filter: string): void {
  const sel = byId<HTMLSelectElement>("modelSelect");
  const f = filter.trim().toLowerCase();
  const matches = f ? allModels.filter((m) => m.toLowerCase().includes(f)) : allModels;
  sel.innerHTML = "";
  for (const m of matches.slice(0, 500)) {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    sel.appendChild(opt);
  }
  sel.size = Math.min(8, Math.max(2, matches.length));
}

function onPickModel(): void {
  const sel = byId<HTMLSelectElement>("modelSelect");
  if (sel.value) byId<HTMLInputElement>("model").value = sel.value;
}

function onProviderChange(): void {
  // Different provider -> different model catalog; hide the stale picker.
  allModels = [];
  byId<HTMLInputElement>("modelFilter").classList.add("hidden");
  byId<HTMLSelectElement>("modelSelect").classList.add("hidden");
  updateKeyHint();
}

function updateKeyHint(): void {
  const spec = getProvider(byId<HTMLSelectElement>("provider").value);
  const key = byId<HTMLInputElement>("apiKey");
  key.disabled = !!spec && !spec.requiresKey;
  key.placeholder = spec && !spec.requiresKey ? "(no key needed)" : "sk-…";
}

function errText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

function setStatus(msg: string, kind: "ok" | "err" | ""): void {
  const el = byId<HTMLParagraphElement>("status");
  el.textContent = msg;
  el.className = "status " + kind;
}

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el as T;
}
