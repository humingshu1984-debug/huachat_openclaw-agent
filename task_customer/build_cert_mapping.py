import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from modules.customer_cert_mapper import CustomerCertMapper
from modules.customer_cert_mapping_reporter import CustomerCertMappingReporter
from modules.customer_coverage_analyzer import CustomerCoverageAnalyzer


def main():
    config_path = os.path.join(PROJECT_ROOT, "task_customer", "config", "config.yaml")
    mapper = CustomerCertMapper(config_path)
    reporter = CustomerCertMappingReporter(config_path)
    analyzer = CustomerCoverageAnalyzer(config_path)

    print("=" * 60)
    print("task_customer 证书关联映射构建")
    print("=" * 60)

    print("\n[阶段1] 生成证书客户关联映射...")
    mapping_bundle = mapper.build_mapping()
    mapping_df = mapping_bundle.get("mapping_df")
    low_conf_df = mapping_bundle.get("low_confidence_df")
    unmatched_df = mapping_bundle.get("unmatched_df")

    paths = mapper.paths
    mapping_file = paths.get("cert_mapping_file")
    os.makedirs(os.path.dirname(mapping_file), exist_ok=True)
    with pd.ExcelWriter(mapping_file, engine="openpyxl") as writer:
        (mapping_df if mapping_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="mapping")
        (low_conf_df if low_conf_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="low_confidence")
        (unmatched_df if unmatched_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="unmatched")
    print(f"✅ 映射文件: {mapping_file}")

    print("\n[阶段2] 计算关联前后认证统计变化...")
    before_dfs = analyzer.analyze(use_cert_mapping=False)
    after_dfs = analyzer.analyze(use_cert_mapping=True)

    print("\n[阶段3] 生成关联效果分析报告...")
    report = reporter.build_report(
        mapping_df=mapping_df if mapping_df is not None else pd.DataFrame(),
        unmatched_df=unmatched_df if unmatched_df is not None else pd.DataFrame(),
        low_conf_df=low_conf_df if low_conf_df is not None else pd.DataFrame(),
        before_region_df=before_dfs.get("region_summary_df"),
        after_region_df=after_dfs.get("region_summary_df"),
    )
    out_files = reporter.save_outputs(report)
    print(f"✅ 报告Excel: {out_files.get('report_excel')}")
    print(f"✅ 报告Markdown: {out_files.get('report_md')}")

    print("\n[阶段4] 结果摘要")
    total = len(mapping_df) + len(unmatched_df)
    print(f"- 证书客户总数: {total}")
    print(f"- 成功关联数: {len(mapping_df)}")
    print(f"- 低置信度数: {len(low_conf_df)}")
    print(f"- 未匹配数: {len(unmatched_df)}")

    print("\n" + "=" * 60)
    print("证书关联映射构建完成！")
    print("=" * 60)


if __name__ == "__main__":
    import pandas as pd
    main()
