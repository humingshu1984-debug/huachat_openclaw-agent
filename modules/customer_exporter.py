import os
import yaml
import pandas as pd
import time


class CustomerExporter:
    def __init__(self, config_path: str):
        with open(config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f) or {}
        self.paths = self.config.get("paths", {}) or {}

    def export_excels(self, dfs: dict):
        out_dir = self.paths.get("output_dir", "outputs/")
        os.makedirs(out_dir, exist_ok=True)

        summary_path = os.path.join(out_dir, "customer_coverage_summary.xlsx")
        detail_path = os.path.join(out_dir, "customer_coverage_detail.xlsx")

        def _with_fallback(path):
            base, ext = os.path.splitext(path)
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            return f"{base}_{timestamp}{ext}"

        def _safe_sheet(name):
            safe = str(name)
            for ch in ['\\', '/', '?', '*', '[', ']', ':']:
                safe = safe.replace(ch, ' ')
            safe = safe.strip() or "Sheet"
            return safe[:31]

        summary_sheets = [
            ("按二级部门汇总", dfs.get("region_summary_df")),
            ("按三级部门汇总", dfs.get("region_level3_summary_df")),
            ("证书类型数量", dfs.get("cert_type_count_df")),
            ("证书类型占比", dfs.get("cert_type_rate_df")),
            ("覆盖x认证", dfs.get("coverage_x_cert_df")),
            ("Top未覆盖客户", dfs.get("top_uncovered_customers_df")),
            ("证书映射记录", dfs.get("cert_mapping_used_df")),
            ("Software认证汇总(大区)", dfs.get("software_region_summary_df")),
            ("Software认证汇总(代表处)", dfs.get("software_subregion_summary_df")),
            ("Software认证客户明细", dfs.get("software_certified_customers_df")),
            ("Software Top未认证客户", dfs.get("software_top_uncertified_customers_df")),
        ]

        try:
            with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
                has_any = False
                for sheet, df in summary_sheets:
                    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
                        continue
                    df.to_excel(writer, index=False, sheet_name=_safe_sheet(sheet))
                    has_any = True
                if not has_any:
                    pd.DataFrame({"info": ["no data"]}).to_excel(writer, index=False, sheet_name="info")
        except PermissionError:
            summary_path = _with_fallback(summary_path)
            with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
                has_any = False
                for sheet, df in summary_sheets:
                    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
                        continue
                    df.to_excel(writer, index=False, sheet_name=_safe_sheet(sheet))
                    has_any = True
                if not has_any:
                    pd.DataFrame({"info": ["no data"]}).to_excel(writer, index=False, sheet_name="info")

        detail_df = dfs.get("customer_detail_df")
        try:
            with pd.ExcelWriter(detail_path, engine="openpyxl") as writer:
                if detail_df is None or not isinstance(detail_df, pd.DataFrame) or detail_df.empty:
                    pd.DataFrame({"info": ["no data"]}).to_excel(writer, index=False, sheet_name="info")
                else:
                    detail_df.to_excel(writer, index=False, sheet_name="客户明细")
        except PermissionError:
            detail_path = _with_fallback(detail_path)
            with pd.ExcelWriter(detail_path, engine="openpyxl") as writer:
                if detail_df is None or not isinstance(detail_df, pd.DataFrame) or detail_df.empty:
                    pd.DataFrame({"info": ["no data"]}).to_excel(writer, index=False, sheet_name="info")
                else:
                    detail_df.to_excel(writer, index=False, sheet_name="客户明细")

        return {"summary_excel": summary_path, "detail_excel": detail_path}
