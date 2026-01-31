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
            return int(values.get(node.id, 0))
        if isinstance(node, ast.Num):
            return int(node.n)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return int(node.value)
        raise FormulaError("Unsupported expression")

    return eval_node(tree)


def validate_data(
    extracted: dict,
    required_metrics: list[str],
    cross_checks: dict[str, str],
    mapping_errors: list[dict],
    mapping_warnings: list[dict],
) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []

    report_month = extracted.get("report_month")
    metrics = extracted.get("metrics", {})
    metrics_list = extracted.get("metrics_list", [])
    cases = extracted.get("cases", [])

    if report_month is None:
        errors.append(
            {
                "type": "error",
                "message": "Месяц отчета не найден в заголовке Word.",
                "source": "document",
                "suggestedFix": "Убедитесь, что заголовок содержит месяц и год.",
            }
        )

    for metric_key in required_metrics:
        if metric_key not in metrics:
            errors.append(
                {
                    "type": "error",
                    "message": f"Не найдена обязательная метрика: {metric_key}.",
                    "source": "summary metrics",
                    "suggestedFix": "Добавьте точную фразу в Word или обновите regex в metrics.yaml.",
                }
            )

    for warn in extracted.get("metric_warnings", []):
        warnings.append(warn)
    for warn in extracted.get("case_warnings", []):
        warnings.append(warn)

    for case in cases:
        if not case.get("registered_date"):
            warnings.append(
                {
                    "type": "warning",
                    "message": f"Нет даты регистрации для дела {case.get('case_id')}",
                    "source": f"table {case.get('table_index')} row {case.get('row_index')}",
                    "suggestedFix": "Убедитесь, что в строке дела указана дата.",
                }
            )
        if not case.get("article"):
            warnings.append(
                {
                    "type": "warning",
                    "message": f"Нет статьи/берене для дела {case.get('case_id')}",
                    "source": f"table {case.get('table_index')} row {case.get('row_index')}",
                    "suggestedFix": "Убедитесь, что в строке дела указана статья.",
                }
            )

    case_ids = [case.get("case_id") for case in cases if case.get("case_id")]
    duplicates = sorted({cid for cid in case_ids if case_ids.count(cid) > 1})
    if duplicates:
        warnings.append(
            {
                "type": "warning",
                "message": f"Повторяющиеся case_id: {', '.join(duplicates)}.",
                "source": "case tables",
                "suggestedFix": "Проверьте дубликаты в Word.",
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
                    "message": f"{out_of_month} дел вне отчетного месяца.",
                    "source": "case tables",
                    "suggestedFix": "Проверьте корректность месяца отчетности.",
                }
            )

    values = {key: data.get("value") for key, data in metrics.items() if isinstance(data.get("value"), int)}
    cross_check_results = []
    for target_key, expression in cross_checks.items():
        expected = int(values.get(target_key, 0))
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
                "contributions": {k: int(values.get(k, 0)) for k in values},
            }
        )
        if error_message:
            errors.append(
                {
                    "type": "error",
                    "message": f"Ошибка формулы для {target_key}: {error_message}.",
                    "source": "crossChecks",
                    "suggestedFix": "Проверьте ключи в формуле cross-check.",
                }
            )
        elif not passed:
            errors.append(
                {
                    "type": "error",
                    "message": f"Cross-check не сошелся для {target_key}: ожидалось {expected}, рассчитано {actual}.",
                    "source": "cross-checks",
                    "suggestedFix": "Проверьте метрики и формулу.",
                }
            )

    errors.extend(mapping_errors)
    warnings.extend(mapping_warnings)

    return {
        "errors": errors,
        "warnings": warnings,
        "cross_checks": cross_check_results,
        "metrics_list": metrics_list,
    }
