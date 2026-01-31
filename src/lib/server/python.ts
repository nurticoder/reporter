import { spawn } from "child_process";

export type PythonResult = {
  stdout: string;
  stderr: string;
  exitCode: number | null;
};

type SpawnError = Error & { code?: string };

function isEnoent(error: unknown): error is SpawnError {
  return Boolean(
    error &&
      typeof error === "object" &&
      "code" in error &&
      (error as SpawnError).code === "ENOENT"
  );
}

async function spawnPython(
  command: string,
  args: string[],
  cwd?: string
): Promise<PythonResult> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd,
      env: {
        ...process.env,
        PYTHONIOENCODING: "utf-8",
        PYTHONUTF8: "1",
      },
      windowsHide: true,
    });

    let stdout = "";
    let stderr = "";

    child.stdout?.on("data", (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr?.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("error", (error) => reject(error));
    child.on("close", (exitCode) => resolve({ stdout, stderr, exitCode }));
  });
}

export async function runPython(
  args: string[],
  cwd?: string
): Promise<PythonResult> {
  const configured =
    process.env.PYTHON_EXECUTABLE ||
    process.env.PYTHON ||
    process.env.PYTHON_BIN ||
    "python";

  try {
    return await spawnPython(configured, args, cwd);
  } catch (error) {
    if (isEnoent(error) && configured !== "py") {
      return await spawnPython("py", args, cwd);
    }

    const details = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to start Python: ${details}`);
  }
}
