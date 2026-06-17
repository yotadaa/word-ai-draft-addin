import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const ports = process.argv.slice(2).length ? process.argv.slice(2) : ["3000", "3001"];
const listeners = await getListeners(ports);
const pids = [...new Set(listeners.map((row) => row.pid).filter(Boolean))];

if (!pids.length) {
  console.log(`Ports ${ports.map((port) => `:${port}`).join(", ")} are free.`);
  process.exit(0);
}

for (const pid of pids) {
  if (String(pid) === String(process.pid)) {
    continue;
  }

  await killPid(pid);
  console.log(`Stopped process ${pid} on add-in dev port.`);
}

async function getListeners(targetPorts) {
  if (process.platform === "win32") {
    const { stdout } = await execFileAsync("netstat", ["-ano"]);

    return stdout
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map(parseWindowsNetstatLine)
      .filter(Boolean)
      .filter((row) => row.state === "LISTENING")
      .filter((row) => targetPorts.some((port) => row.localAddress.endsWith(`:${port}`)));
  }

  const { stdout } = await execFileAsync("sh", [
    "-c",
    `lsof -nP -iTCP:${targetPorts.join(",")} -sTCP:LISTEN -t || true`
  ]);

  return stdout
    .split(/\r?\n/)
    .map((pid) => pid.trim())
    .filter(Boolean)
    .map((pid) => ({ pid }));
}

async function killPid(pid) {
  if (process.platform === "win32") {
    await execFileAsync("taskkill", ["/PID", String(pid), "/F", "/T"]);
    return;
  }

  await execFileAsync("kill", [String(pid)]);
}

function parseWindowsNetstatLine(line) {
  const parts = line.split(/\s+/);

  if (parts.length < 5 || !["TCP", "UDP"].includes(parts[0])) {
    return null;
  }

  return {
    protocol: parts[0],
    localAddress: parts[1],
    state: parts[3] ?? "",
    pid: parts[4] ?? ""
  };
}
