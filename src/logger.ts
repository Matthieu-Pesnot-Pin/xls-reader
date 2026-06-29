import fs from "fs";
import path from "path";

let currentPrefix = "";

export function logToFile(message: string, prefix: string = ""): void {
  const defaultLogDir =
    process.platform === "win32"
      ? "C:\\var\\log\\xls-reader"
      : "/var/log/xls-reader";
  const logDir = process.env.MCP_LOG_DIR ?? defaultLogDir;
  const logFile = path.join(logDir, "server.log");

  const now = new Date().toLocaleString("fr-FR", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });

  const effectivePrefix = prefix || currentPrefix;
  const label = effectivePrefix ? ` [${effectivePrefix}]` : "";
  const logEntry = `[${now}] [PID ${process.pid}]${label} ${message}\n`;

  try {
    if (!fs.existsSync(logDir)) {
      fs.mkdirSync(logDir, { recursive: true });
    }
    fs.appendFileSync(logFile, logEntry);
  } catch {
    process.stderr.write("Critical: Failed to write to log file\n");
  }
}

export function setupLogging(prefix: string): void {
  currentPrefix = prefix;

  if (prefix === "MASTER") {
    console.log = (...args: unknown[]): void => {
      console.error(...args);
    };
  }

  const originalError = console.error;

  console.error = (...args: unknown[]): void => {
    const message = args
      .map((arg) => (typeof arg === "object" ? JSON.stringify(arg, null, 2) : arg))
      .join(" ");
    logToFile(message, prefix);
    originalError.apply(console, args);
  };

  logToFile("--- Logging system initialized ---", prefix);
}
