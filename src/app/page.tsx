"use client";

import { useMemo, useState } from "react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";

const ACCEPTED_WORD = ".docx";
const ACCEPTED_EXCEL = ".xlsx";

type Issue = {
  type: "error" | "warning";
  message: string;
  source?: string | null;
  suggestedFix?: string | null;
};

type MetricRow = {
  key: string;
  value: number;
  sourceSnippet: string;
  sourcePointer?: string | null;
};

type ArticleRow = {
  article: string;
  women_u18: number;
  women_ge18: number;
  women_total: number;
  stopped: number;
  new: number;
  total_cases: number;
};

type CrossCheck = {
  key: string;
  formula: string;
  expected: number;
  actual: number;
  pass: boolean;
};

type UpdatePreview = {
  sheet: string;
  cell: string;
  rowLabel: string;
  oldValue: string | number | null;
  newValue: string | number | null;
  kind: string;
};

type AnalysisReport = {
  reportMonth: string;
  extractedMetrics: MetricRow[];
  articleBreakdown: ArticleRow[];
  crossChecks: CrossCheck[];
  errors: Issue[];
  warnings: Issue[];
  updatePreview: UpdatePreview[];
  validationSkipped?: boolean;
};

type AnalysisResponse = {
  status: "ok" | "validation_error";
  report: AnalysisReport;
};

type ErrorResponse = {
  error: string;
  details?: string;
  report?: AnalysisReport;
};

function Spinner() {
  return (
    <div className="h-5 w-5 animate-spin rounded-full border-2 border-slate-200 border-t-slate-900" />
  );
}

function FileCard({
  title,
  description,
  accept,
  file,
  onChange,
}: {
  title: string;
  description: string;
  accept: string;
  file: File | null;
  onChange: (file: File | null) => void;
}) {
  return (
    <Card className="border-slate-200 bg-white/85">
      <CardHeader>
        <CardTitle className="text-base text-slate-900">{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-3">
        <Input
          type="file"
          accept={accept}
          onChange={(event) => {
            const selected = event.currentTarget.files?.[0] ?? null;
            onChange(selected);
          }}
        />
        <div className="text-xs text-slate-500">
          {file ? `Selected: ${file.name}` : "No file selected yet."}
        </div>
      </CardContent>
    </Card>
  );
}

export default function Home() {
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [analysis, setAnalysis] = useState<AnalysisReport | null>(null);
  const [status, setStatus] = useState<"idle" | "analyzing" | "generating">(
    "idle"
  );
  const [apiError, setApiError] = useState<string | null>(null);
  const [skipValidation, setSkipValidation] = useState(false);

  const canAnalyze = wordFile && excelFile && status === "idle";

  const validationIssues = useMemo(() => {
    if (!analysis) return [];
    return [
      ...analysis.errors.map((issue) => ({ ...issue, type: "error" as const })),
      ...analysis.warnings.map((issue) => ({ ...issue, type: "warning" as const })),
    ];
  }, [analysis]);

  const canGenerate =
    wordFile &&
    excelFile &&
    status === "idle" &&
    (skipValidation || (analysis && analysis.errors.length === 0));

  const handleAnalyze = async () => {
    if (!wordFile || !excelFile) return;
    setStatus("analyzing");
    setApiError(null);
    setAnalysis(null);
    const formData = new FormData();
    formData.append("word", wordFile);
    formData.append("excel", excelFile);
    formData.append("skipValidation", skipValidation ? "true" : "false");

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        body: formData,
      });
      let data: AnalysisResponse | ErrorResponse | null = null;
      try {
        data = (await response.json()) as AnalysisResponse | ErrorResponse;
      } catch (parseError) {
        data = null;
      }

      const hasReport = Boolean(data && "report" in data && data.report);
      if (hasReport && data && "report" in data && data.report) {
        setAnalysis(data.report);
      }

      if (!response.ok) {
        const baseMessage =
          data && "error" in data && data.error
            ? data.error
            : hasReport
              ? "Validation issues detected. Review the report below."
              : "Analyze failed. Please check the server logs.";
        const details =
          data && "details" in data && data.details
            ? ` Details: ${data.details}`
            : "";
        setApiError(`${baseMessage}${details}`);
      }
    } catch (error) {
      setApiError("Analyze failed. Please check the server logs.");
    } finally {
      setStatus("idle");
    }
  };

  const handleGenerate = async () => {
    if (!wordFile || !excelFile) return;
    setStatus("generating");
    setApiError(null);
    const formData = new FormData();
    formData.append("word", wordFile);
    formData.append("excel", excelFile);
    formData.append("skipValidation", skipValidation ? "true" : "false");

    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });
      if (!response.ok) {
        let data: AnalysisResponse | ErrorResponse | null = null;
        try {
          data = (await response.json()) as AnalysisResponse | ErrorResponse;
        } catch (parseError) {
          data = null;
        }

        const hasReport = Boolean(data && "report" in data && data.report);
        if (hasReport && data && "report" in data && data.report) {
          setAnalysis(data.report);
        }

        const baseMessage =
          data && "error" in data && data.error
            ? data.error
            : hasReport
              ? "Generation blocked. Review the validation report."
              : "Generate failed. Please check the server logs.";
        const details =
          data && "details" in data && data.details
            ? ` Details: ${data.details}`
            : "";
        setApiError(`${baseMessage}${details}`);
        return;
      }

      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition");
      const filenameMatch = disposition?.match(/filename="(.+?)"/);
      const filename = filenameMatch?.[1] ?? "report_updated.xlsx";
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = filename;
      anchor.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      setApiError("Generate failed. Please check the server logs.");
    } finally {
      setStatus("idle");
    }
  };

  return (
    <div className="min-h-screen px-6 py-12">
      <div className="mx-auto flex w-full max-w-6xl flex-col gap-8">
        <header className="space-y-3">
          <Badge variant="default" className="w-fit border-slate-300 text-slate-600">
            Local-first report updater
          </Badge>
          <h1 className="heading-serif text-3xl font-semibold text-slate-950 sm:text-4xl">
            Update monthly Excel reports from Word documents with deterministic checks.
          </h1>
          <p className="max-w-3xl text-base text-slate-600">
            Upload the current month Word report and the previous month Excel base.
            The analyzer extracts metrics, validates everything, and only allows
            Excel generation when all checks pass.
          </p>
        </header>

        <section className="grid gap-6 lg:grid-cols-2">
          <FileCard
            title="Current month Word report (.docx)"
            description="Source of metrics and case table details."
            accept={ACCEPTED_WORD}
            file={wordFile}
            onChange={setWordFile}
          />
          <FileCard
            title="Previous month Excel report (.xlsx)"
            description="The base file that will be updated."
            accept={ACCEPTED_EXCEL}
            file={excelFile}
            onChange={setExcelFile}
          />
        </section>

        <section className="flex flex-wrap items-center gap-4">
          <Button onClick={handleAnalyze} disabled={!canAnalyze}>
            {status === "analyzing" ? (
              <span className="flex items-center gap-2">
                <Spinner />
                Analyzing...
              </span>
            ) : (
              "Analyze"
            )}
          </Button>
          <Button
            variant="secondary"
            onClick={handleGenerate}
            disabled={!canGenerate}
          >
            {status === "generating" ? (
              <span className="flex items-center gap-2">
                <Spinner />
                Generating...
              </span>
            ) : (
              "Generate & Download Excel"
            )}
          </Button>
          {analysis && (
            <span className="text-sm text-slate-500">
              Report month: <span className="font-medium text-slate-900">{analysis.reportMonth}</span>
            </span>
          )}
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <Input
              type="checkbox"
              checked={skipValidation}
              onChange={(event) => setSkipValidation(event.currentTarget.checked)}
              className="h-4 w-4"
            />
            Skip validation (unsafe)
          </label>
        </section>

        {apiError && (
          <Card className="border-amber-200 bg-amber-50/70">
            <CardContent className="py-4 text-sm text-amber-800">
              {apiError}
            </CardContent>
          </Card>
        )}

        {analysis && (
          <section className="grid gap-6">
            <Card>
              <CardHeader>
                <CardTitle>Validation results</CardTitle>
                <CardDescription>
                  Issues block generation. Warnings highlight edge cases for review.
                </CardDescription>
              </CardHeader>
              <CardContent>
                {validationIssues.length === 0 ? (
                  <div className="text-sm text-emerald-700">
                    All validations passed. You can generate the Excel file.
                  </div>
                ) : (
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Issue Type</TableHead>
                        <TableHead>Description</TableHead>
                        <TableHead>Source</TableHead>
                        <TableHead>Suggested Fix</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {validationIssues.map((issue, index) => (
                        <TableRow key={`${issue.type}-${index}`}>
                          <TableCell className="capitalize">{issue.type}</TableCell>
                          <TableCell>{issue.message}</TableCell>
                          <TableCell>{issue.source ?? "—"}</TableCell>
                          <TableCell>{issue.suggestedFix ?? "—"}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                )}
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle>Extracted summary metrics</CardTitle>
                <CardDescription>Metric name, value, and source snippet.</CardDescription>
              </CardHeader>
              <CardContent>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Metric</TableHead>
                      <TableHead>Value</TableHead>
                      <TableHead>Source snippet</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {analysis.extractedMetrics.map((metric) => (
                      <TableRow key={metric.key}>
                        <TableCell className="font-medium text-slate-900">
                          {metric.key}
                        </TableCell>
                        <TableCell>{metric.value}</TableCell>
                        <TableCell className="text-slate-600">
                          {metric.sourceSnippet}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle>Article breakdown</CardTitle>
                <CardDescription>
                  Article -&gt; women &lt; 18, women &gt;= 18, women total, stopped, new, totals.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Article</TableHead>
                      <TableHead>Women &lt; 18</TableHead>
                      <TableHead>Women &gt;= 18</TableHead>
                      <TableHead>Women total</TableHead>
                      <TableHead>Stopped</TableHead>
                      <TableHead>New</TableHead>
                      <TableHead>Totals</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {analysis.articleBreakdown.map((row) => (
                      <TableRow key={row.article}>
                        <TableCell className="font-medium text-slate-900">
                          {row.article}
                        </TableCell>
                        <TableCell>{row.women_u18}</TableCell>
                        <TableCell>{row.women_ge18}</TableCell>
                        <TableCell>{row.women_total}</TableCell>
                        <TableCell>{row.stopped}</TableCell>
                        <TableCell>{row.new}</TableCell>
                        <TableCell>{row.total_cases}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle>Cross-checks</CardTitle>
                <CardDescription>Consistency checks across metrics.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-3">
                {analysis.crossChecks.map((check) => (
                  <div
                    key={check.key}
                    className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 bg-slate-50/70 p-3"
                  >
                    <div>
                      <div className="font-medium text-slate-900">{check.key}</div>
                      <div className="text-xs text-slate-500">{check.formula}</div>
                    </div>
                    <div className="text-sm text-slate-600">
                      expected {check.expected}, actual {check.actual}
                    </div>
                    <Badge variant={check.pass ? "success" : "attention"}>
                      {check.pass ? "OK" : "Needs attention"}
                    </Badge>
                  </div>
                ))}
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle>Preview diff</CardTitle>
                <CardDescription>
                  Old Excel value -&gt; proposed new value -&gt; cell reference.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Sheet</TableHead>
                      <TableHead>Row label</TableHead>
                      <TableHead>Cell</TableHead>
                      <TableHead>Old</TableHead>
                      <TableHead>New</TableHead>
                      <TableHead>Kind</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {analysis.updatePreview.map((row, index) => (
                      <TableRow key={`${row.sheet}-${row.cell}-${index}`}>
                        <TableCell>{row.sheet}</TableCell>
                        <TableCell>{row.rowLabel}</TableCell>
                        <TableCell>{row.cell}</TableCell>
                        <TableCell>{row.oldValue ?? "—"}</TableCell>
                        <TableCell>{row.newValue ?? "—"}</TableCell>
                        <TableCell>{row.kind}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </CardContent>
            </Card>
          </section>
        )}
      </div>
    </div>
  );
}

