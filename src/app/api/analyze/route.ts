import { NextResponse } from "next/server";
import { fileSync } from "tmp";
import fs from "fs/promises";
import path from "path";
import { runPython } from "@/lib/server/python";

export const runtime = "nodejs";

async function writeTempFile(file: File, postfix: string) {
  const tmpFile = fileSync({ postfix, discardDescriptor: true });
  const buffer = Buffer.from(await file.arrayBuffer());
  await fs.writeFile(tmpFile.name, buffer);
  return tmpFile;
}

export async function POST(request: Request) {
  const formData = await request.formData();
  const word = formData.get("word");
  const excel = formData.get("excel");
  const skipValidation = formData.get("skipValidation") === "true";

  if (!(word instanceof File) || !(excel instanceof File)) {
    return NextResponse.json(
      { error: "Both word and excel files are required." },
      { status: 400 }
    );
  }

  const wordTmp = await writeTempFile(word, ".docx");
  const excelTmp = await writeTempFile(excel, ".xlsx");

  try {
    const scriptPath = path.join(process.cwd(), "processor", "main.py");
    const args = [
      scriptPath,
      "analyze",
      "--word",
      wordTmp.name,
      "--excel",
      excelTmp.name,
    ];
    if (skipValidation) {
      args.push("--skip-validation");
    }

    const result = await runPython(args, process.cwd());

    if (!result.stdout.trim()) {
      return NextResponse.json(
        {
          error: "Python worker did not return output.",
          details: result.stderr,
        },
        { status: 500 }
      );
    }

    const payload = JSON.parse(result.stdout);
    if (payload.status !== "ok") {
      return NextResponse.json(payload, { status: 422 });
    }

    return NextResponse.json(payload, { status: 200 });
  } catch (error) {
    return NextResponse.json(
      {
        error: "Failed to analyze report.",
        details: String(error),
      },
      { status: 500 }
    );
  } finally {
    wordTmp.removeCallback();
    excelTmp.removeCallback();
  }
}

