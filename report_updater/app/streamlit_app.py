from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.apply_excel import apply_updates, plan_updates
from core.config_loader import PROJECT_ROOT, load_yaml
from core.diff_template import diff_template
from core.excel_inspect import inspect_excel
from core.extract_docx import extract_docx
from core.validate import validate_data

OUTPUT_DIR = PROJECT_ROOT / "output"
INPUT_DIR = OUTPUT_DIR / "inputs"


def load_i18n(lang_code: str) -> dict[str, Any]:
    name = f"i18n_{lang_code}.yaml"
    try:
        return load_yaml(name)
    except FileNotFoundError:
        return load_yaml("i18n_ru.yaml")


def save_upload(uploaded, suffix: str) -> Path:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = Path(uploaded.name).name
    if not safe_name.lower().endswith(suffix):
        safe_name = f"{safe_name}{suffix}"
    path = INPUT_DIR / safe_name
    path.write_bytes(uploaded.getbuffer())
    return path


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def render_table(title: str, rows: list[dict]) -> None:
    st.subheader(title)
    if rows:
        st.dataframe(rows, use_container_width=True)
    else:
        st.info("—")


def main() -> None:
    st.set_page_config(page_title="Локальный апдейтер отчетов МВД", layout="wide")

    lang_choice = st.sidebar.selectbox("Язык / Тил", ["Русский", "Кыргызча"])
    lang_code = "ru" if lang_choice == "Русский" else "kg"
    i18n = load_i18n(lang_code)

    st.title(i18n["app_title"])

    docx_file = st.file_uploader(i18n["upload_docx"], type=["docx"])
    prev_file = st.file_uploader(i18n["upload_prev_xlsx"], type=["xlsx"])
    template_file = st.file_uploader(i18n["upload_template_xlsx"], type=["xlsx"])

    skip_validation = st.checkbox(i18n["skip_validation"])
    confirm_skip = False
    if skip_validation:
        st.warning(i18n["skip_warning"])
        confirm_skip = st.checkbox(i18n["skip_confirm"])

    analyze_clicked = st.button(i18n["analyze_button"])
    if analyze_clicked:
        if not docx_file or not prev_file:
            st.error(i18n["missing_inputs"])
        else:
            docx_path = save_upload(docx_file, ".docx")
            prev_path = save_upload(prev_file, ".xlsx")
            template_path = save_upload(template_file, ".xlsx") if template_file else None

            metrics_cfg = load_yaml("metrics.yaml")
            if isinstance(metrics_cfg, dict) and "metrics" in metrics_cfg:
                metrics_cfg = metrics_cfg["metrics"]
            required_cfg = load_yaml("metrics_required.yaml")
            if isinstance(required_cfg, dict):
                required_cfg = required_cfg.get("required", [])
            cross_checks = load_yaml("cross_checks.yaml")
            if isinstance(cross_checks, dict) and "cross_checks" in cross_checks:
                cross_checks = cross_checks["cross_checks"]
            excel_map = load_yaml("excel_map.yaml")
            article_map = load_yaml("article_map.yaml")

            extracted = extract_docx(str(docx_path), metrics_cfg)
            updates, map_errors, map_warnings, map_debug = plan_updates(
                str(prev_path), extracted["metrics"], extracted["article_breakdown"], excel_map, article_map
            )
            if not updates:
                map_errors.append(
                    {
                        "type": "error",
                        "message": "Не найдено ни одной цели записи в Excel.",
                        "source": "mapping",
                        "suggestedFix": "Проверьте excel_map.yaml и структуру файла.",
                    }
                )

            validation = validate_data(
                extracted,
                required_cfg,
                cross_checks,
                map_errors,
                map_warnings,
            )

            prev_inspect = inspect_excel(str(prev_path))
            template_inspect = inspect_excel(str(template_path)) if template_path else None
            template_diff = diff_template(str(prev_path), str(template_path)) if template_path else []

            analysis = {
                "docx_path": str(docx_path),
                "prev_path": str(prev_path),
                "template_path": str(template_path) if template_path else None,
                "extracted": extracted,
                "validation": validation,
                "updates": updates,
                "mapping_debug": map_debug,
                "prev_inspect": prev_inspect,
                "template_inspect": template_inspect,
                "template_diff": template_diff,
            }
            st.session_state["analysis"] = analysis

            write_json(OUTPUT_DIR / "extracted_metrics.json", extracted)
            write_json(OUTPUT_DIR / "validation_report.json", validation)
            write_json(OUTPUT_DIR / "mapping_debug.json", {"updates": updates, "debug": map_debug})

            st.success(i18n["analysis_done"])

    analysis = st.session_state.get("analysis")
    if analysis:
        st.subheader(i18n["excel_overview"])
        st.markdown(f"**{i18n['excel_prev_title']}**")
        st.json(analysis["prev_inspect"], expanded=False)
        if analysis.get("template_inspect"):
            st.markdown(f"**{i18n['excel_template_title']}**")
            st.json(analysis["template_inspect"], expanded=False)
            st.subheader(i18n["diff_template_title"])
            if analysis["template_diff"]:
                st.dataframe(analysis["template_diff"], use_container_width=True)
            else:
                st.info(i18n["diff_template_empty"])

        report_month = analysis["extracted"].get("report_month")
        if report_month:
            st.markdown(f"**{i18n['report_month']}:** {report_month.get('label')}")

        st.subheader(i18n["validation_title"])
        errors = analysis["validation"]["errors"]
        warnings = analysis["validation"]["warnings"]
        if errors:
            st.error(i18n["blocked"])
        else:
            st.success(i18n["no_errors"])
        render_table(i18n["errors_title"], errors)
        render_table(i18n["warnings_title"], warnings)

        render_table(i18n["cross_checks_title"], analysis["validation"]["cross_checks"])
        render_table(i18n["metrics_title"], analysis["validation"]["metrics_list"])
        render_table(i18n["articles_title"], analysis["extracted"]["article_breakdown"])
        render_table(i18n["preview_title"], analysis["updates"])

        if st.button(i18n["generate_button"]):
            if errors and not (skip_validation and confirm_skip):
                st.error(i18n["blocked"])
            else:
                base_name = Path(analysis["prev_path"]).stem
                output_name = f"{base_name}_generated.xlsx"
                output_path = OUTPUT_DIR / output_name
                apply_updates(analysis["prev_path"], str(output_path), analysis["updates"])
                st.success(i18n["generated"])
                st.info(i18n["download_ready"])
                with output_path.open("rb") as handle:
                    st.download_button(
                        label=i18n["generate_button"],
                        data=handle.read(),
                        file_name=output_name,
                    )


if __name__ == "__main__":
    main()
