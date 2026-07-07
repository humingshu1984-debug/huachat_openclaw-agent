import os
import yaml
import pandas as pd


class CustomerCertMappingReporter:
    def __init__(self, config_path: str):
        with open(config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f) or {}
        self.paths = self.config.get("paths", {}) or {}
        self.analysis_cfg = self.config.get("analysis", {}) or {}

    def _build_overall_summary(self, mapping_df, unmatched_df, low_conf_df):
        total = int(len(mapping_df) + len(unmatched_df))
        matched = int(len(mapping_df))
        unmatched = int(len(unmatched_df))
        low_conf = int(len(low_conf_df))
        return pd.DataFrame([{
            "证书客户总量": total,
            "成功关联量": matched,
            "未关联量": unmatched,
            "成功关联率": round(matched / total * 100, 2) if total else 0.0,
            "低置信度量": low_conf,
        }])

    def _build_region_summary(self, mapping_df, unmatched_df):
        matched = mapping_df.groupby("mapped_region")["cert_gsp_id"].nunique().reset_index(name="成功关联量") if not mapping_df.empty else pd.DataFrame(columns=["mapped_region", "成功关联量"])
        matched = matched.rename(columns={"mapped_region": "大区"})

        unmatched = unmatched_df.groupby("source_region")["cert_gsp_id"].nunique().reset_index(name="未关联量") if not unmatched_df.empty else pd.DataFrame(columns=["source_region", "未关联量"])
        unmatched = unmatched.rename(columns={"source_region": "大区"})

        region = matched.merge(unmatched, on="大区", how="outer").fillna(0)
        if region.empty:
            return pd.DataFrame(columns=["大区", "成功关联量", "未关联量", "关联总量", "关联率"])
        region["成功关联量"] = region["成功关联量"].astype(int)
        region["未关联量"] = region["未关联量"].astype(int)
        region["关联总量"] = region["成功关联量"] + region["未关联量"]
        region["关联率"] = region.apply(lambda r: round(r["成功关联量"] / r["关联总量"] * 100, 2) if r["关联总量"] else 0.0, axis=1)
        return region.sort_values(by=["关联率", "成功关联量", "大区"], ascending=[False, False, True]).reset_index(drop=True)

    def _build_before_after(self, before_df, after_df):
        if before_df is None or before_df.empty:
            before_df = pd.DataFrame(columns=["region_l2", "certified_count", "certified_ratio"])
        if after_df is None or after_df.empty:
            after_df = pd.DataFrame(columns=["region_l2", "certified_count", "certified_ratio"])

        before = before_df[["region_l2", "certified_count", "certified_ratio"]].copy() if "region_l2" in before_df.columns else pd.DataFrame(columns=["region_l2", "certified_count", "certified_ratio"])
        after = after_df[["region_l2", "certified_count", "certified_ratio"]].copy() if "region_l2" in after_df.columns else pd.DataFrame(columns=["region_l2", "certified_count", "certified_ratio"])
        before = before.rename(columns={"region_l2": "大区", "certified_count": "关联前认证通过个数", "certified_ratio": "关联前认证通过比例"})
        after = after.rename(columns={"region_l2": "大区", "certified_count": "关联后认证通过个数", "certified_ratio": "关联后认证通过比例"})
        merged = before.merge(after, on="大区", how="outer").fillna(0)
        if merged.empty:
            return pd.DataFrame(columns=["大区"])
        merged["认证通过个数变化"] = merged["关联后认证通过个数"] - merged["关联前认证通过个数"]
        merged["认证通过比例变化"] = (merged["关联后认证通过比例"] - merged["关联前认证通过比例"]).round(2)
        return merged.sort_values(by=["认证通过个数变化", "认证通过比例变化", "大区"], ascending=[False, False, True]).reset_index(drop=True)

    def _build_recommendations(self, overall_summary_df, region_summary_df):
        thresholds = self.analysis_cfg.get("cert_mapping_report_thresholds", {}) or {}
        overall_unmatched_threshold = float(thresholds.get("overall_unmatched_ratio_for_search", 0.35))
        region_unmatched_threshold = float(thresholds.get("region_unmatched_ratio_for_search", 0.40))
        low_conf_threshold = float(thresholds.get("low_confidence_ratio_for_search", 0.25))

        overall = overall_summary_df.iloc[0].to_dict() if overall_summary_df is not None and not overall_summary_df.empty else {}
        total = float(overall.get("证书客户总量", 0) or 0)
        unmatched = float(overall.get("未关联量", 0) or 0)
        low_conf = float(overall.get("低置信度量", 0) or 0)
        unmatched_ratio = unmatched / total if total else 0.0
        low_conf_ratio = low_conf / total if total else 0.0

        overall_advice = "建议先不引入 Google 搜索"
        if unmatched_ratio >= overall_unmatched_threshold or low_conf_ratio >= low_conf_threshold:
            overall_advice = "建议评估引入 Google 搜索等第二阶段手段"

        region_advice_rows = []
        if region_summary_df is not None and not region_summary_df.empty:
            for _, row in region_summary_df.iterrows():
                ratio = float(row.get("未关联量", 0)) / float(row.get("关联总量", 1) or 1)
                advice = "建议先不引入 Google 搜索"
                if ratio >= region_unmatched_threshold:
                    advice = "建议该大区优先进入第二阶段补充搜索"
                region_advice_rows.append({
                    "大区": row.get("大区", ""),
                    "未关联占比": round(ratio * 100, 2),
                    "建议": advice,
                })

        return overall_advice, pd.DataFrame(region_advice_rows)

    def build_report(self, mapping_df, unmatched_df, low_conf_df, before_region_df, after_region_df):
        overall_summary_df = self._build_overall_summary(mapping_df, unmatched_df, low_conf_df)
        region_summary_df = self._build_region_summary(mapping_df, unmatched_df)
        before_after_df = self._build_before_after(before_region_df, after_region_df)
        overall_advice, region_advice_df = self._build_recommendations(overall_summary_df, region_summary_df)

        added_certified_df = before_after_df[before_after_df.get("认证通过个数变化", 0) > 0].copy() if not before_after_df.empty else pd.DataFrame(columns=["大区"])

        return {
            "overall_summary_df": overall_summary_df,
            "region_summary_df": region_summary_df,
            "before_after_df": before_after_df,
            "added_certified_df": added_certified_df,
            "unmatched_df": unmatched_df if unmatched_df is not None else pd.DataFrame(),
            "low_confidence_df": low_conf_df if low_conf_df is not None else pd.DataFrame(),
            "overall_advice": overall_advice,
            "region_advice_df": region_advice_df,
        }

    def save_outputs(self, report: dict):
        excel_path = self.paths.get("cert_mapping_report_excel")
        md_path = self.paths.get("cert_mapping_report_md")
        os.makedirs(os.path.dirname(excel_path), exist_ok=True)

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            report.get("overall_summary_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="整体汇总")
            report.get("region_summary_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="分大区汇总")
            report.get("before_after_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="关联前后对比")
            report.get("added_certified_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="新增认证客户")
            report.get("unmatched_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="未关联客户")
            report.get("low_confidence_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="低置信度候选")
            report.get("region_advice_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="建议结论")

        lines = []
        lines.append("# 数据关联执行结果分析报告")
        lines.append("")
        overall = report.get("overall_summary_df", pd.DataFrame())
        if not overall.empty:
            row = overall.iloc[0]
            lines.append("## 整体关联效果总结")
            lines.append(f"- 证书客户总量：{int(row['证书客户总量'])}")
            lines.append(f"- 成功关联量：{int(row['成功关联量'])}")
            lines.append(f"- 未关联量：{int(row['未关联量'])}")
            lines.append(f"- 成功关联率：{row['成功关联率']:.2f}%")
            lines.append(f"- 低置信度量：{int(row['低置信度量'])}")
            lines.append(f"- 整体建议：{report.get('overall_advice', '')}")
            lines.append("")

        region_df = report.get("region_summary_df", pd.DataFrame())
        if not region_df.empty:
            lines.append("## 分大区关联效果概览")
            for _, row in region_df.iterrows():
                lines.append(f"- {row['大区']}: 成功关联 {int(row['成功关联量'])}，未关联 {int(row['未关联量'])}，关联率 {row['关联率']:.2f}%")
            lines.append("")

        before_after = report.get("before_after_df", pd.DataFrame())
        if not before_after.empty:
            lines.append("## 关联前后认证统计变化")
            for _, row in before_after.iterrows():
                lines.append(
                    f"- {row['大区']}: 认证通过个数 {int(row['关联前认证通过个数'])} -> {int(row['关联后认证通过个数'])}，"
                    f"比例 {row['关联前认证通过比例']:.2f}% -> {row['关联后认证通过比例']:.2f}%"
                )
            lines.append("")

        advice_df = report.get("region_advice_df", pd.DataFrame())
        if not advice_df.empty:
            lines.append("## 分大区第二阶段建议")
            for _, row in advice_df.iterrows():
                lines.append(f"- {row['大区']}: 未关联占比 {row['未关联占比']:.2f}%，{row['建议']}")
            lines.append("")

        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines).strip() + "\n")

        return {"report_excel": excel_path, "report_md": md_path}
