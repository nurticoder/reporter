from __future__ import annotations

import ast
from datetime import datetime
from typing import Any


class FormulaError(Exception):
    pass


def safe_eval(expression: str, values: dict[str, int]) -> int:
    tree = ast.parse(expression, mode="eval")

    def eval_node(node: ast.AST) -> int:
        if isinstance(node, ast.Expression):
            return eval_node(node.body)
        if isinstance(node, ast.BinOp):
            left = eval_node(node.left)
            right = eval_node(node.right)
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            raise FormulaError("Unsupported operator")
        if isinstance(node, ast.UnaryOp):
            operand = eval_node(node.operand)
            if isinstance(node.op, ast.USub):
                return -operand
            if isinstance(node.op, ast.UAdd):
                return operand
            raise FormulaError("Unsupported unary operator")
        if isinstance(node, ast.Name):
            return values.get(node.id, 0)
        if isinstance(node, ast.Num):
            return int(node.n)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return int(node.value)
        raise FormulaError("Unsupported expression")

    return eval_node(tree)


def validate_report(
    report_month: dict | None,
    metrics: dict,
    metrics_list: list[dict],
    cases: list[dict],
    required_metrics: list[str],
    cross_checks: dict,
    article_breakdown: list[dict],
    article_map: dict,
    issues: list[dict],
) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []
    unmapped_metrics: list[str] = []

    if report_month is None:
        errors.append(
            {
                "type": "error",
                "message": "Report month could not be extracted from the Word title.",
                "source": "document",
                "suggestedFix": "Ensure the title contains a month/year like 'за ноябрь 2025 г.'.",
            }
        )

    for metric_key in required_metrics:
        if metric_key not in metrics:
            unmapped_metrics.append(metric_key)
            errors.append(
                {
                    "type": "error",
                    "message": f"Missing required metric: {metric_key}.",
                    "source": "summary metrics",
                    "suggestedFix": "Add the exact phrase to the Word report or update the metric dictionary.",
                }
            )

    for key, data in metrics.items():
        value = data.get("value")
        if not isinstance(value, int):
            errors.append(
                {
                    "type": "error",
                    "message": f"Metric {key} is non-numeric.",
                    "source": data.get("sourcePointer"),
                    "suggestedFix": "Ensure the metric value is a number.",
                }
            )

    duplicate_meta = None
    for issue in issues:
        if issue.get("code") == "duplicate_case_ids":
            duplicate_meta = issue
            continue
        if issue.get("type") == "error":
            errors.append(issue)
        else:
            warnings.append(issue)

    if duplicate_meta:
        duplicate_count = int(duplicate_meta.get("duplicateCount") or 0)
        total_cases = int(duplicate_meta.get("totalCases") or 0)
        ratio = duplicate_count / max(1, total_cases)
        warnings.append(duplicate_meta)
        if ratio > 0.1:
            errors.append(
                {
                    "type": "error",
                    "message": "Duplicate case IDs exceed 10% of rows.",
                    "source": "case tables",
                    "suggestedFix": "Remove duplicate case rows before retrying.",
                }
            )

    missing_date_count = sum(1 for case in cases if not case.get("registered_date"))
    missing_article_count = sum(1 for case in cases if not case.get("article_base"))
    total_case_count = len(cases)
    if total_case_count:
        if missing_date_count / total_case_count > 0.5:
            errors.append(
                {
                    "type": "error",
                    "message": "Missing registered dates for more than 50% of cases.",
                    "source": "case tables",
                    "suggestedFix": "Ensure each case row contains a registration date.",
                }
            )
        if missing_article_count / total_case_count > 0.5:
            errors.append(
                {
                    "type": "error",
                    "message": "Missing article codes for more than 50% of cases.",
                    "source": "case tables",
                    "suggestedFix": "Ensure each case row contains a 'бер' or 'ст.' reference.",
                }
            )

    if report_month:
        out_of_month = 0
        for case in cases:
            date_str = case.get("registered_date")
            if not date_str:
                continue
            try:
                parsed = datetime.fromisoformat(date_str)
            except ValueError:
                continue
            if parsed.year != report_month["year"] or parsed.month != report_month["month"]:
                out_of_month += 1
        if out_of_month:
            warnings.append(
                {
                    "type": "warning",
                    "message": f"{out_of_month} cases fall outside the report month.",
                    "source": "case tables",
                    "suggestedFix": "Confirm whether to include non-report-month cases.",
                }
            )

    cross_check_results = []
    values = {key: data.get("value") for key, data in metrics.items() if isinstance(data.get("value"), int)}
    for target_key, expression in cross_checks.items():
        expected = values.get(target_key, 0)
        actual = None
        passed = False
        error_message = None
        try:
            actual = safe_eval(expression, values)
            passed = expected == actual
        except FormulaError as exc:
            error_message = str(exc)
        cross_check_results.append(
            {
                "key": target_key,
                "formula": expression,
                "expected": expected,
                "actual": actual,
                "pass": passed,
            }
        )
        if error_message:
            errors.append(
                {
                    "type": "error",
                    "message": f"Cross-check formula error for {target_key}: {error_message}.",
                    "source": "crossChecks.json",
                    "suggestedFix": "Verify metric keys in crossChecks.json.",
                }
            )
        elif not passed:
            errors.append(
                {
                    "type": "error",
                    "message": f"Cross-check failed for {target_key}: expected {expected}, calculated {actual}.",
                    "source": "cross-checks",
                    "suggestedFix": "Verify the metric values and adjust if needed.",
                }
            )

    allowed_articles = set(article_map.get("allowedArticles", []))
    if allowed_articles:
        for row in article_breakdown:
            article = row.get("article")
            if article and article not in allowed_articles:
                warnings.append(
                    {
                        "type": "warning",
                        "message": f"Unmapped article {article}.",
                        "source": "case tables",
                        "suggestedFix": "Add the article to config/articleMap.json allowedArticles.",
                    }
                )

    return {
        "errors": errors,
        "warnings": warnings,
        "cross_checks": cross_check_results,
        "metrics_list": metrics_list,
        "unmapped_metrics": unmapped_metrics,
    }

