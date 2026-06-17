import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";

dotenv.config({ path: ".env.local" });
dotenv.config();

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
  const envLocal = path.resolve(".env.local");
  const env = path.resolve(".env");

  report(fs.existsSync(envLocal) || fs.existsSync(env), "env", [
    `.env.local: ${fs.existsSync(envLocal) ? "found" : "missing"}`,
    `.env: ${fs.existsSync(env) ? "found" : "missing"}`,
    `AI_BASE_URL: ${baseUrl}`,
    `AI_MODEL: ${model}`,
    `AI_API_KEY/OPENAI_API_KEY: ${apiKey ? "set" : "missing"}`
  ]);
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

function normalizeBaseUrl(value) {
  return value.replace(/\/+$/, "");
}

function formatProviderError(body) {
  return body?.error?.message ?? body?.message ?? "No chat completion content returned.";
}
