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
            if node.id not in values:
                raise FormulaError(f"Missing metric in formula: {node.id}")
            return values[node.id]
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

    issues_by_type = [issue for issue in issues]
    for issue in issues_by_type:
        if issue.get("type") == "error":
            errors.append(issue)
        else:
            warnings.append(issue)

    duplicate_ids = []
    seen_ids = set()
    for case in cases:
        case_id = case.get("case_id")
        if not case_id:
            continue
        if case_id in seen_ids:
            duplicate_ids.append(case_id)
        else:
            seen_ids.add(case_id)
    if duplicate_ids:
        errors.append(
            {
                "type": "error",
                "message": f"Duplicate case IDs detected: {', '.join(sorted(set(duplicate_ids)))}.",
                "source": "case tables",
                "suggestedFix": "Remove duplicate case rows before retrying.",
            }
        )

    date_months = set()
    for case in cases:
        date_str = case.get("registered_date")
        if not date_str:
            continue
        try:
            parsed = datetime.fromisoformat(date_str)
            date_months.add((parsed.year, parsed.month))
        except ValueError:
            continue
    if len(date_months) > 1:
        warnings.append(
            {
                "type": "warning",
                "message": "Multiple registration months detected in case tables.",
                "source": "case tables",
                "suggestedFix": "Confirm all rows belong to the same reporting month.",
            }
        )

    cross_check_results = []
    values = {key: data.get("value") for key, data in metrics.items() if isinstance(data.get("value"), int)}
    for target_key, expression in cross_checks.items():
        expected = values.get(target_key)
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
        elif expected is None:
            errors.append(
                {
                    "type": "error",
                    "message": f"Cross-check target missing metric: {target_key}.",
                    "source": "summary metrics",
                    "suggestedFix": "Ensure the target metric is extracted.",
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
                errors.append(
                    {
                        "type": "error",
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

