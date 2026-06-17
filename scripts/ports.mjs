import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const ports = process.argv.slice(2).length ? process.argv.slice(2) : ["3000", "3001"];

const rows = await getPortRows(ports);

if (!rows.length) {
  console.log(`No listeners found on ${ports.map((port) => `:${port}`).join(", ")}.`);
  process.exit(0);
}

for (const row of rows) {
  console.log(`${row.protocol} ${row.localAddress} ${row.state} pid=${row.pid}`);
}

async function getPortRows(targetPorts) {
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
    `lsof -nP -iTCP:${targetPorts.join(",")} -sTCP:LISTEN || true`
  ]);

  return stdout
    .split(/\r?\n/)
    .slice(1)
    .map((line) => line.trim().split(/\s+/))
    .filter((parts) => parts.length >= 9)
    .map((parts) => ({
      protocol: "TCP",
      localAddress: parts[8],
      state: "LISTENING",
      pid: parts[1]
    }));
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
