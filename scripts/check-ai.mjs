import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptDir, "..");
const envLocalPath = path.join(projectRoot, ".env.local");
const envPath = path.join(projectRoot, ".env");

const loadedEnvLocal = dotenv.config({ path: envLocalPath });
const loadedEnv = dotenv.config({ path: envPath });

const apiPort = Number(process.env.API_PORT ?? 3001);
const baseUrl = normalizeBaseUrl(process.env.AI_BASE_URL ?? "http://localhost:20128/v1");
const apiKey = process.env.AI_API_KEY ?? process.env.OPENAI_API_KEY ?? "";
const model = process.env.AI_MODEL?.trim() || "cx/gpt-5.5";

const checks = [];

await checkEnvFiles();
await checkProviderModels();
await checkProviderChat();
await checkAddinApiHealth();

const failed = checks.some((check) => check.status === "fail");

console.log("");
console.log(failed ? "AI check failed." : "AI check passed.");
process.exitCode = failed ? 1 : 0;

async function checkEnvFiles() {
  const envHints = [
    `cwd: ${process.cwd()}`,
    `project root: ${projectRoot}`,
    `.env.local: ${fs.existsSync(envLocalPath) ? `found (${envLocalPath})` : "missing"}`,
    `.env: ${fs.existsSync(envPath) ? `found (${envPath})` : "missing"}`,
    `.env.local keys: ${formatParsedKeys(loadedEnvLocal.parsed)}`,
    `.env keys: ${formatParsedKeys(loadedEnv.parsed)}`,
    `process AI_API_KEY: ${process.env.AI_API_KEY ? "set" : "missing"}`,
    `process OPENAI_API_KEY: ${process.env.OPENAI_API_KEY ? "set" : "missing"}`,
    `AI_BASE_URL: ${baseUrl}`,
    `AI_MODEL: ${model}`,
    `AI_API_KEY/OPENAI_API_KEY: ${apiKey ? `set (${redactKey(apiKey)})` : "missing"}`
  ];
  const envWarnings = getEnvWarnings();

  report(Boolean(apiKey), "env", [...envHints, ...envWarnings]);
}

async function checkProviderModels() {
  try {
    const response = await fetchWithTimeout(`${baseUrl}/models`, {
      headers: authHeaders()
    });
    const body = await response.json().catch(() => null);
    const modelIds = Array.isArray(body?.data) ? body.data.map((item) => item.id).filter(Boolean) : [];

    report(response.ok, "provider /models", [
      `HTTP ${response.status}`,
      modelIds.length ? `models: ${modelIds.slice(0, 8).join(", ")}` : "models: none returned"
    ]);
  } catch (error) {
    report(false, "provider /models", [
      error instanceof Error ? error.message : String(error),
      `Make sure the OpenAI-compatible provider is running on ${baseUrl}.`
    ]);
  }
}

async function checkProviderChat() {
  try {
    const response = await fetchWithTimeout(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...authHeaders()
      },
      body: JSON.stringify({
        model,
        temperature: 0,
        messages: [{ role: "user", content: "Balas hanya kata: ok" }]
      })
    });
    const body = await response.json().catch(() => null);
    const content = body?.choices?.[0]?.message?.content;

    report(response.ok && Boolean(content), "provider chat", [
      `HTTP ${response.status}`,
      content ? `response: ${String(content).slice(0, 80)}` : formatProviderError(body)
    ]);
  } catch (error) {
    report(false, "provider chat", [error instanceof Error ? error.message : String(error)]);
  }
}

async function checkAddinApiHealth() {
  try {
    const response = await fetchWithTimeout(`http://127.0.0.1:${apiPort}/api/health`);
    const body = await response.json().catch(() => null);

    report(response.ok && Boolean(body?.ok), "addin API /api/health", [
      `HTTP ${response.status}`,
      body ? JSON.stringify(body) : "No JSON body",
      `Run npm run dev in another terminal before this check if this fails.`
    ]);
  } catch (error) {
    report(false, "addin API /api/health", [
      error instanceof Error ? error.message : String(error),
      "Run npm run dev and keep it open."
    ]);
  }
}

function fetchWithTimeout(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10000);

  return fetch(url, {
    ...options,
    signal: controller.signal
  }).finally(() => clearTimeout(timeout));
}

function authHeaders() {
  return apiKey ? { Authorization: `Bearer ${apiKey}` } : {};
}

function report(ok, name, lines) {
  const status = ok ? "pass" : "fail";
  checks.push({ name, status });
  console.log(`${ok ? "PASS" : "FAIL"} ${name}`);
  lines.forEach((line) => console.log(`  ${line}`));
}

function formatParsedKeys(parsed) {
  const keys = Object.keys(parsed ?? {});

  return keys.length ? keys.join(", ") : "none";
}

function getEnvWarnings() {
  const warnings = [];
  const envFiles = [
    { label: ".env.local", path: envLocalPath },
    { label: ".env", path: envPath }
  ];

  envFiles.forEach((file) => {
    if (!fs.existsSync(file.path)) {
      return;
    }

    const buffer = fs.readFileSync(file.path);
    const text = buffer.toString("utf8");

    if (buffer.includes(0)) {
      warnings.push(`${file.label}: file looks like UTF-16/Unicode. Re-save it as UTF-8.`);
    }

    if (/AI_API_KEY\s*:/i.test(text)) {
      warnings.push(`${file.label}: use AI_API_KEY=... not AI_API_KEY: ...`);
    }

    if (/^\s*(?:export\s+)?AI_API_KEY\s*=/im.test(text) && !loadedEnvLocal.parsed?.AI_API_KEY && !loadedEnv.parsed?.AI_API_KEY) {
      warnings.push(`${file.label}: AI_API_KEY line exists but dotenv did not parse it. Check quotes, encoding, or hidden extension.`);
    }

    if (/AI_API_KEY\s*=/i.test(text) && /AI_API_KEY\s*=\s*["']?\s*["']?\s*(?:\r?\n|$)/i.test(text)) {
      warnings.push(`${file.label}: AI_API_KEY appears empty.`);
    }
  });

  if (!fs.existsSync(envLocalPath) && fs.existsSync(path.join(projectRoot, ".env.local.txt"))) {
    warnings.push("Found .env.local.txt; rename it to .env.local.");
  }

  return warnings;
}

function redactKey(value) {
  const trimmed = String(value).trim();

  if (trimmed.length <= 8) {
    return "short value";
  }

  return `${trimmed.slice(0, 5)}...${trimmed.slice(-4)}`;
}

function normalizeBaseUrl(value) {
  return value.replace(/\/+$/, "");
}

function formatProviderError(body) {
  return body?.error?.message ?? body?.message ?? "No chat completion content returned.";
}
