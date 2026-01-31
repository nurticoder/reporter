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
  const outTmp = fileSync({ postfix: ".xlsx", discardDescriptor: true });

  try {
    const scriptPath = path.join(process.cwd(), "processor", "main.py");
    const args = [
      scriptPath,
      "generate",
      "--word",
      wordTmp.name,
      "--excel",
      excelTmp.name,
      "--out",
      outTmp.name,
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

    const fileBuffer = await fs.readFile(outTmp.name);
    const reportMonth = payload.report?.reportMonth ?? "report";
    const filename = `report_updated_${reportMonth}.xlsx`;

    return new NextResponse(fileBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (error) {
    return NextResponse.json(
      {
        error: "Failed to generate report.",
        details: String(error),
      },
      { status: 500 }
    );
  } finally {
    wordTmp.removeCallback();
    excelTmp.removeCallback();
    outTmp.removeCallback();
  }
}

