from __future__ import annotations

import argparse
import json
from pathlib import Path

from config_loader import (
    load_article_map,
    load_cross_checks,
    load_excel_map,
    load_metric_dictionary,
    load_required_metrics,
)
from extract_word import extract_word_data
from update_excel import apply_updates, plan_updates, sha256_file
from validate import validate_report


def build_report(
    report_month,
    metrics_list,
    article_breakdown,
    cross_checks,
    errors,
    warnings,
    update_preview,
    unmapped_metrics,
    validation_skipped,
):
    suggested_fixes = []
    for issue in errors + warnings:
        if issue.get("suggestedFix"):
            suggested_fixes.append(issue["suggestedFix"])
    return {
        "reportMonth": report_month.get("label") if report_month else "unknown",
        "extractedMetrics": metrics_list,
        "articleBreakdown": article_breakdown,
        "crossChecks": cross_checks,
        "errors": errors,
        "warnings": warnings,
        "unmappedMetrics": unmapped_metrics,
        "suggestedFixes": suggested_fixes,
        "updatePreview": update_preview,
        "validationSkipped": validation_skipped,
    }


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)

    analyze_parser = subparsers.add_parser("analyze")
    analyze_parser.add_argument("--word", required=True)
    analyze_parser.add_argument("--excel", required=True)
    analyze_parser.add_argument("--skip-validation", action="store_true")

    generate_parser = subparsers.add_parser("generate")
    generate_parser.add_argument("--word", required=True)
    generate_parser.add_argument("--excel", required=True)
    generate_parser.add_argument("--out", required=True)
    generate_parser.add_argument("--skip-validation", action="store_true")

    args = parser.parse_args()

    metric_dictionary = load_metric_dictionary()
    required_metrics = load_required_metrics()
    excel_map = load_excel_map()
    article_map = load_article_map()
    cross_checks = load_cross_checks()

    word_data = extract_word_data(args.word, metric_dictionary)

    validation = validate_report(
        report_month=word_data["report_month"],
        metrics=word_data["metrics"],
        metrics_list=word_data["metrics_list"],
        cases=word_data["cases"],
        required_metrics=required_metrics,
        cross_checks=cross_checks,
        article_breakdown=word_data["article_breakdown"],
        article_map=article_map,
        issues=word_data["metric_issues"] + word_data["case_issues"] + word_data["warnings"],
    )

    plan_updates_list, plan_errors = plan_updates(
        args.excel,
        word_data["metrics"],
        word_data["article_breakdown"],
        excel_map,
        article_map,
    )

    if plan_errors:
        validation["errors"].extend(plan_errors)
    if not plan_updates_list:
        validation["errors"].append(
            {
                "type": "error",
                "message": "No Excel targets resolved from the mapping configuration.",
                "source": "excelMap.json",
                "suggestedFix": "Verify sheet names, row labels, and column headers match the Excel template.",
            }
        )

    report = build_report(
        report_month=word_data["report_month"],
        metrics_list=validation["metrics_list"],
        article_breakdown=word_data["article_breakdown"],
        cross_checks=validation["cross_checks"],
        errors=validation["errors"],
        warnings=validation["warnings"],
        update_preview=[
            {
                "sheet": item["sheet"],
                "cell": item["cell"],
                "rowLabel": item["rowLabel"],
                "oldValue": item["oldValue"],
                "newValue": item["newValue"],
                "kind": item["kind"],
            }
            for item in plan_updates_list
        ],
        unmapped_metrics=validation.get("unmapped_metrics", []),
        validation_skipped=args.skip_validation,
    )

    if validation["errors"] and not args.skip_validation:
        print(
            json.dumps(
                {
                    "status": "validation_error",
                    "report": report,
                },
                ensure_ascii=False,
            )
        )
        return

    if args.command == "generate":
        word_hash = sha256_file(args.word)
        excel_hash = sha256_file(args.excel)
        counts = {
            "metrics": len(word_data["metrics"].keys()),
            "cases": len(word_data["cases"]),
            "articles": len(word_data["article_breakdown"]),
        }
        summary = "Validation passed. Updates applied."
        apply_updates(
            args.excel,
            args.out,
            plan_updates_list,
            report_month=word_data["report_month"],
            word_hash=word_hash,
            excel_hash=excel_hash,
            counts=counts,
            summary=summary,
        )

    print(
        json.dumps(
            {
                "status": "ok",
                "report": report,
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()

