import copy
import os
import sys
from datetime import datetime

import pandas as pd
import yaml

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from modules.customer_cert_mapper import CustomerCertMapper


def _normalize(value):
    return "".join(str(value).replace("\u3000", " ").split()).strip()


def _pick_col(df, candidates, required=True, label=""):
    if df is None or df.empty:
        if required:
            raise ValueError(f"数据为空，无法识别字段：{label}")
        return None

    norm_map = {_normalize(c): c for c in df.columns}
    for cand in candidates or []:
        key = _normalize(cand)
        if key in norm_map:
            return norm_map[key]

    if required:
        cols_preview = ", ".join([str(c) for c in list(df.columns)[:40]])
        raise ValueError(f"缺少必要字段：{label}。候选={candidates}。实际列={cols_preview}")
    return None


def _abs_path(root, path_value):
    s = str(path_value or "").strip()
    if not s:
        return ""
    if os.path.isabs(s):
        return s
    return os.path.join(root, s)


def _load_yaml(path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _save_yaml(obj, path):
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(obj, f, allow_unicode=True, sort_keys=False)


def _snapshot_path(output_dir):
    return os.path.join(output_dir, "dhia_cert_seen_snapshot.xlsx")


def _delta_path(output_dir):
    return os.path.join(output_dir, "dhia_cert_delta.xlsx")


def _review_todo_path(output_dir):
    return os.path.join(output_dir, "dhia_cert_incremental_review_todo.xlsx")


def _load_snapshot(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame(columns=["cert_gsp_id", "cert_company_name"])
    try:
        df = pd.read_excel(path, sheet_name="seen_gsp_id")
    except Exception:
        df = pd.read_excel(path)
    df = df.copy()
    if "cert_gsp_id" not in df.columns:
        return pd.DataFrame(columns=["cert_gsp_id", "cert_company_name"])
    if "cert_company_name" not in df.columns:
        df["cert_company_name"] = ""
    df["cert_gsp_id"] = df["cert_gsp_id"].astype(str).str.strip()
    df["cert_company_name"] = df["cert_company_name"].astype(str).str.strip()
    df = df[df["cert_gsp_id"].str.len() > 0].drop_duplicates(subset=["cert_gsp_id"])
    return df[["cert_gsp_id", "cert_company_name"]].copy()


def _save_snapshot(df, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="seen_gsp_id")


def _norm_df_text(df, cols):
    if df is None or df.empty:
        return df
    for c in cols:
        if c in df.columns:
            df[c] = df[c].astype(str).fillna("").str.strip()
    return df


def _build_review_todo(
    new_preview,
    mapping_df,
    low_conf_df,
    unmatched_df,
    renamed_df,
):
    new_preview = new_preview.copy() if new_preview is not None else pd.DataFrame()
    mapping_df = mapping_df.copy() if mapping_df is not None else pd.DataFrame()
    low_conf_df = low_conf_df.copy() if low_conf_df is not None else pd.DataFrame()
    unmatched_df = unmatched_df.copy() if unmatched_df is not None else pd.DataFrame()
    renamed_df = renamed_df.copy() if renamed_df is not None else pd.DataFrame()

    new_preview = _norm_df_text(
        new_preview,
        ["cert_gsp_id", "cert_company_name", "cert_region", "cert_sub_region", "cert_country", "cert_type_list"],
    )
    mapping_df = _norm_df_text(
        mapping_df,
        [
            "cert_gsp_id",
            "cert_company_name",
            "cert_region",
            "cert_sub_region",
            "mapped_customer_id",
            "mapped_customer_name",
            "mapped_region",
            "mapped_sub_region",
            "match_basis",
            "match_status",
        ],
    )
    low_conf_df = _norm_df_text(
        low_conf_df,
        [
            "cert_gsp_id",
            "cert_company_name",
            "cert_region",
            "cert_sub_region",
            "mapped_customer_id",
            "mapped_customer_name",
            "mapped_region",
            "mapped_sub_region",
            "match_basis",
            "match_status",
        ],
    )
    unmatched_df = _norm_df_text(
        unmatched_df,
        ["cert_gsp_id", "cert_company_name", "cert_region", "cert_sub_region"],
    )
    renamed_df = _norm_df_text(renamed_df, ["cert_gsp_id", "old_company_name", "new_company_name"])

    strong_basis = {"账号直接匹配", "名称标准化后完全一致"}

    def first_row(df, cert_id):
        if df is None or df.empty or "cert_gsp_id" not in df.columns:
            return None
        rows = df[df["cert_gsp_id"] == cert_id]
        if rows.empty:
            return None
        return rows.iloc[0].to_dict()

    must_rows = []
    suggest_rows = []
    for _, base in (new_preview or pd.DataFrame()).iterrows():
        cert_id = str(base.get("cert_gsp_id", "")).strip()
        if not cert_id:
            continue

        row = {
            "cert_gsp_id": cert_id,
            "cert_company_name": str(base.get("cert_company_name", "")).strip(),
            "cert_region": str(base.get("cert_region", "")).strip(),
            "cert_sub_region": str(base.get("cert_sub_region", "")).strip(),
            "cert_country": str(base.get("cert_country", "")).strip(),
            "cert_type_list": str(base.get("cert_type_list", "")).strip(),
            "cert_row_count": int(base.get("cert_row_count", 0) or 0),
        }

        u = first_row(unmatched_df, cert_id)
        if u is not None:
            row.update({"review_level": "必须审核", "review_reason": "未匹配"})
            must_rows.append(row)
            continue

        l = first_row(low_conf_df, cert_id)
        if l is not None:
            row.update(
                {
                    "mapped_customer_id": str(l.get("mapped_customer_id", "")).strip(),
                    "mapped_customer_name": str(l.get("mapped_customer_name", "")).strip(),
                    "mapped_region": str(l.get("mapped_region", "")).strip(),
                    "mapped_sub_region": str(l.get("mapped_sub_region", "")).strip(),
                    "match_basis": str(l.get("match_basis", "")).strip(),
                    "match_score": float(l.get("match_score", 0) or 0),
                    "match_status": str(l.get("match_status", "")).strip(),
                    "review_level": "必须审核",
                    "review_reason": "低置信度",
                }
            )
            must_rows.append(row)
            continue

        m = first_row(mapping_df, cert_id)
        if m is not None:
            row.update(
                {
                    "mapped_customer_id": str(m.get("mapped_customer_id", "")).strip(),
                    "mapped_customer_name": str(m.get("mapped_customer_name", "")).strip(),
                    "mapped_region": str(m.get("mapped_region", "")).strip(),
                    "mapped_sub_region": str(m.get("mapped_sub_region", "")).strip(),
                    "match_basis": str(m.get("match_basis", "")).strip(),
                    "match_score": float(m.get("match_score", 0) or 0),
                    "match_status": str(m.get("match_status", "")).strip(),
                }
            )
            basis = row.get("match_basis", "")
            if basis in strong_basis:
                continue
            row.update({"review_level": "建议抽查", "review_reason": "弱匹配规则"})
            suggest_rows.append(row)
            continue

        row.update({"review_level": "必须审核", "review_reason": "缺少结果行"})
        must_rows.append(row)

    must_df = pd.DataFrame(must_rows)
    suggest_df = pd.DataFrame(suggest_rows)
    return must_df, suggest_df, renamed_df


def _save_review_todo(
    output_path,
    summary_rows,
    must_df,
    suggest_df,
    renamed_df,
):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="summary")
        (must_df if must_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="must_review")
        (suggest_df if suggest_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="suggest_review")
        (renamed_df if renamed_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="renamed_company")


def _build_cert_dim(cert_df, cert_cols):
    id_col = _pick_col(cert_df, cert_cols.get("customer_id_candidates"), label="客户ID(证书清单)")
    name_col = _pick_col(cert_df, cert_cols.get("customer_name_candidates"), required=False, label="客户名称(证书清单)")
    cert_type_col = _pick_col(cert_df, cert_cols.get("cert_type_candidates"), required=False, label="证书类型(证书清单)")
    region_col = _pick_col(cert_df, cert_cols.get("region_l2_candidates"), required=False, label="证书客户大区")
    sub_region_col = _pick_col(cert_df, cert_cols.get("region_l3_candidates"), required=False, label="证书客户代表处")
    country_col = _pick_col(cert_df, cert_cols.get("country_candidates"), required=False, label="证书客户国家")

    dim = pd.DataFrame()
    dim["cert_gsp_id"] = cert_df[id_col].astype(str).str.strip()
    dim["cert_company_name"] = cert_df[name_col].astype(str).str.strip() if name_col else ""
    dim["cert_type"] = cert_df[cert_type_col].astype(str).str.strip() if cert_type_col else ""
    dim["cert_region"] = cert_df[region_col].astype(str).str.strip() if region_col else ""
    dim["cert_sub_region"] = cert_df[sub_region_col].astype(str).str.strip() if sub_region_col else ""
    dim["cert_country"] = cert_df[country_col].astype(str).str.strip() if country_col else ""
    dim = dim[dim["cert_gsp_id"].str.len() > 0].copy()
    return dim, id_col


def _merge_mapping_files(main_path, incremental_path, new_ids):
    new_ids = {str(x).strip() for x in (new_ids or []) if str(x).strip()}
    if not new_ids:
        return {"added_mapping": 0, "added_low_conf": 0, "added_unmatched": 0}

    inc_mapping = pd.read_excel(incremental_path, sheet_name="mapping")
    inc_low = pd.read_excel(incremental_path, sheet_name="low_confidence")
    inc_unmatched = pd.read_excel(incremental_path, sheet_name="unmatched")

    for df in [inc_mapping, inc_low, inc_unmatched]:
        if df is None or df.empty:
            continue
        if "cert_gsp_id" in df.columns:
            df["cert_gsp_id"] = df["cert_gsp_id"].astype(str).str.strip()

    inc_mapping = inc_mapping[inc_mapping.get("cert_gsp_id", pd.Series(dtype=str)).isin(list(new_ids))].copy()
    inc_low = inc_low[inc_low.get("cert_gsp_id", pd.Series(dtype=str)).isin(list(new_ids))].copy()
    inc_unmatched = inc_unmatched[inc_unmatched.get("cert_gsp_id", pd.Series(dtype=str)).isin(list(new_ids))].copy()

    if os.path.exists(main_path):
        xls = pd.ExcelFile(main_path)
        sheets = {name: pd.read_excel(main_path, sheet_name=name) for name in xls.sheet_names}
    else:
        sheets = {}

    main_mapping = sheets.get("mapping", pd.DataFrame()).copy()
    main_low = sheets.get("low_confidence", pd.DataFrame()).copy()
    main_unmatched = sheets.get("unmatched", pd.DataFrame()).copy()

    existing_ids = set()
    for df in [main_mapping, main_low, main_unmatched]:
        if df is not None and not df.empty and "cert_gsp_id" in df.columns:
            existing_ids |= set(df["cert_gsp_id"].astype(str).str.strip().tolist())

    to_add_ids = new_ids - existing_ids
    if not to_add_ids:
        return {"added_mapping": 0, "added_low_conf": 0, "added_unmatched": 0}

    inc_mapping_add = inc_mapping[inc_mapping["cert_gsp_id"].isin(list(to_add_ids))].copy() if not inc_mapping.empty else pd.DataFrame()
    inc_low_add = inc_low[inc_low["cert_gsp_id"].isin(list(to_add_ids))].copy() if not inc_low.empty else pd.DataFrame()
    inc_unmatched_add = inc_unmatched[inc_unmatched["cert_gsp_id"].isin(list(to_add_ids))].copy() if not inc_unmatched.empty else pd.DataFrame()

    merged_mapping = pd.concat([main_mapping, inc_mapping_add], ignore_index=True) if not main_mapping.empty else inc_mapping_add
    merged_low = pd.concat([main_low, inc_low_add], ignore_index=True) if not main_low.empty else inc_low_add
    merged_unmatched = pd.concat([main_unmatched, inc_unmatched_add], ignore_index=True) if not main_unmatched.empty else inc_unmatched_add

    sheets["mapping"] = merged_mapping
    sheets["low_confidence"] = merged_low
    sheets["unmatched"] = merged_unmatched

    os.makedirs(os.path.dirname(main_path), exist_ok=True)
    with pd.ExcelWriter(main_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            (df if df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name=str(name)[:31])

    return {"added_mapping": len(inc_mapping_add), "added_low_conf": len(inc_low_add), "added_unmatched": len(inc_unmatched_add)}


def run_incremental(config_path: str):
    if not config_path:
        config_path = os.path.join(PROJECT_ROOT, "task_customer", "config", "config.yaml")
    if not os.path.isabs(config_path):
        config_path = os.path.join(PROJECT_ROOT, config_path)

    cfg = _load_yaml(config_path)
    paths = cfg.get("paths", {}) or {}
    columns = cfg.get("columns", {}) or {}
    cert_cols = columns.get("cert", {}) or {}

    output_dir = _abs_path(PROJECT_ROOT, paths.get("output_dir") or os.path.join("task_customer", "outputs"))
    cert_file = _abs_path(PROJECT_ROOT, paths.get("cert_file"))
    main_mapping_file = _abs_path(PROJECT_ROOT, paths.get("cert_mapping_file"))

    snap_path = _snapshot_path(output_dir)
    delta_path = _delta_path(output_dir)
    review_path = _review_todo_path(output_dir)
    incremental_cert_path = os.path.join(output_dir, "dhia_certificates_incremental.xlsx")
    incremental_mapping_path = os.path.join(output_dir, "cert_customer_mapping_incremental.xlsx")
    incremental_config_path = os.path.join(output_dir, "config_cert_incremental.yaml")

    print("=" * 60)
    print("task_customer 证书增量映射")
    print("=" * 60)
    print(f"- cert_file: {cert_file}")
    print(f"- snapshot: {snap_path}")
    print(f"- mapping_file: {main_mapping_file}")

    cert_df = pd.read_excel(cert_file)
    cert_dim, id_col = _build_cert_dim(cert_df, cert_cols)

    snapshot_exists = os.path.exists(snap_path)
    snapshot_df = _load_snapshot(snap_path)
    snapshot_ids = set(snapshot_df["cert_gsp_id"].astype(str).str.strip().tolist())
    current_ids = set(cert_dim["cert_gsp_id"].astype(str).str.strip().tolist())

    new_ids = sorted(list(current_ids - snapshot_ids))

    renamed_df = pd.DataFrame(columns=["cert_gsp_id", "old_company_name", "new_company_name"])
    if not snapshot_df.empty and "cert_company_name" in snapshot_df.columns:
        cur_name = (
            cert_dim[["cert_gsp_id", "cert_company_name"]]
            .dropna()
            .assign(cert_gsp_id=lambda d: d["cert_gsp_id"].astype(str).str.strip())
        )
        cur_name = cur_name[cur_name["cert_gsp_id"].str.len() > 0].copy()
        cur_name = cur_name.sort_values(by=["cert_gsp_id", "cert_company_name"]).drop_duplicates(subset=["cert_gsp_id"], keep="last")
        snap_name = snapshot_df.copy()
        merged = snap_name.merge(cur_name, on="cert_gsp_id", how="inner", suffixes=("_old", "_new"))
        merged["old_norm"] = merged["cert_company_name_old"].astype(str).str.strip().str.lower()
        merged["new_norm"] = merged["cert_company_name_new"].astype(str).str.strip().str.lower()
        renamed = merged[(merged["old_norm"] != merged["new_norm"]) & (merged["new_norm"].str.len() > 0)].copy()
        if not renamed.empty:
            renamed_df = renamed.rename(columns={"cert_company_name_old": "old_company_name", "cert_company_name_new": "new_company_name"})[
                ["cert_gsp_id", "old_company_name", "new_company_name"]
            ].copy()

    new_preview = cert_dim[cert_dim["cert_gsp_id"].isin(new_ids)].copy()
    if not new_preview.empty:
        new_preview = (
            new_preview.groupby("cert_gsp_id")
            .agg(
                cert_company_name=("cert_company_name", "first"),
                cert_region=("cert_region", "first"),
                cert_sub_region=("cert_sub_region", "first"),
                cert_country=("cert_country", "first"),
                cert_type_list=("cert_type", lambda xs: "; ".join(sorted(set([str(x).strip() for x in xs if str(x).strip()])))),
                cert_row_count=("cert_type", "size"),
            )
            .reset_index()
        )

    os.makedirs(output_dir, exist_ok=True)
    mode_label = "初始化快照" if not snapshot_exists else "增量识别"
    summary_rows = [
        {"指标": "模式", "值": mode_label},
        {"指标": "当前证书GSP ID数", "值": len(current_ids)},
        {"指标": "历史已见GSP ID数", "值": len(snapshot_ids)},
        {"指标": "新增GSP ID数", "值": (0 if not snapshot_exists else len(new_ids))},
        {"指标": "公司名变化GSP ID数", "值": (0 if not snapshot_exists else len(renamed_df))},
        {"指标": "生成时间", "值": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ]
    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(delta_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="summary")
        (new_preview if new_preview is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="new_gsp_id")
        (renamed_df if renamed_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="renamed_company")

    print(f"✅ 增量识别结果: {delta_path}")
    if not snapshot_exists:
        latest_snapshot = (
            cert_dim.groupby("cert_gsp_id")
            .agg(cert_company_name=("cert_company_name", "first"))
            .reset_index()
            .sort_values(by=["cert_gsp_id"])
        )
        _save_snapshot(latest_snapshot, snap_path)
        print("✅ 首次运行：已初始化快照，不做映射合并。")
        return {"mode": "init_snapshot", "new_gsp_id_count": 0, "renamed_count": 0}

    print(f"- 新增GSP ID: {len(new_ids)}")
    if len(new_ids) == 0:
        latest_snapshot = (
            cert_dim.groupby("cert_gsp_id")
            .agg(cert_company_name=("cert_company_name", "first"))
            .reset_index()
            .sort_values(by=["cert_gsp_id"])
        )
        _save_snapshot(latest_snapshot, snap_path)
        print("✅ 本次无新增证书客户ID，已刷新快照，无需映射。")
        if renamed_df is not None and not renamed_df.empty:
            must_df = pd.DataFrame(columns=["cert_gsp_id", "review_level", "review_reason"])
            suggest_df = pd.DataFrame(columns=["cert_gsp_id", "review_level", "review_reason"])
            summary_rows2 = list(summary_rows) + [
                {"指标": "必须审核条数", "值": 0},
                {"指标": "建议抽查条数", "值": 0},
                {"指标": "输出review_todo", "值": review_path},
            ]
            _save_review_todo(review_path, summary_rows2, must_df, suggest_df, renamed_df)
            print(f"✅ 审核清单: {review_path}")
        return {"mode": "no_new_id", "new_gsp_id_count": 0, "renamed_count": len(renamed_df)}

    cert_subset = cert_df[cert_df[id_col].astype(str).str.strip().isin(new_ids)].copy()
    with pd.ExcelWriter(incremental_cert_path, engine="openpyxl") as writer:
        cert_subset.to_excel(writer, index=False, sheet_name="certificates")
    print(f"✅ 增量证书文件: {incremental_cert_path}")

    cfg2 = copy.deepcopy(cfg)
    cfg2.setdefault("paths", {})
    cfg2["paths"]["cert_file"] = incremental_cert_path
    cfg2["paths"]["cert_mapping_file"] = incremental_mapping_path
    _save_yaml(cfg2, incremental_config_path)

    print("\n[阶段1] 生成新增证书的映射...")
    mapper = CustomerCertMapper(incremental_config_path)
    bundle = mapper.build_mapping()
    mapping_df = bundle.get("mapping_df") if bundle else pd.DataFrame()
    low_conf_df = bundle.get("low_confidence_df") if bundle else pd.DataFrame()
    unmatched_df = bundle.get("unmatched_df") if bundle else pd.DataFrame()

    with pd.ExcelWriter(incremental_mapping_path, engine="openpyxl") as writer:
        (mapping_df if mapping_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="mapping")
        (low_conf_df if low_conf_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="low_confidence")
        (unmatched_df if unmatched_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="unmatched")
    print(f"✅ 增量映射文件: {incremental_mapping_path}")

    must_df, suggest_df, renamed_out_df = _build_review_todo(new_preview, mapping_df, low_conf_df, unmatched_df, renamed_df)
    summary_rows2 = list(summary_rows) + [
        {"指标": "必须审核条数", "值": len(must_df) if must_df is not None else 0},
        {"指标": "建议抽查条数", "值": len(suggest_df) if suggest_df is not None else 0},
        {"指标": "输出review_todo", "值": review_path},
    ]
    _save_review_todo(review_path, summary_rows2, must_df, suggest_df, renamed_out_df)
    print(f"✅ 审核清单: {review_path}")

    print("\n[阶段2] 合并增量映射到主映射文件...")
    backup_path = ""
    if main_mapping_file and os.path.exists(main_mapping_file):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(output_dir, f"cert_customer_mapping.backup_before_incremental_{stamp}.xlsx")
        if not os.path.exists(backup_path):
            import shutil

            shutil.copyfile(main_mapping_file, backup_path)
    merge_stats = _merge_mapping_files(main_mapping_file, incremental_mapping_path, new_ids)
    print(f"✅ 合并完成: {main_mapping_file}")
    if backup_path:
        print(f"- 备份: {backup_path}")
    print(f"- 新增 mapping: {merge_stats.get('added_mapping')}")
    print(f"- 新增 low_confidence: {merge_stats.get('added_low_conf')}")
    print(f"- 新增 unmatched: {merge_stats.get('added_unmatched')}")

    latest_snapshot = (
        cert_dim.groupby("cert_gsp_id")
        .agg(cert_company_name=("cert_company_name", "first"))
        .reset_index()
        .sort_values(by=["cert_gsp_id"])
    )
    _save_snapshot(latest_snapshot, snap_path)
    print(f"✅ 已刷新快照: {snap_path}")

    print("\n" + "=" * 60)
    print("证书增量映射完成！")
    print("=" * 60)
    return {
        "mode": "merged",
        "new_gsp_id_count": len(new_ids),
        "renamed_count": len(renamed_df),
        "merge_stats": merge_stats,
        "backup_path": backup_path,
        "delta_path": delta_path,
        "snapshot_path": snap_path,
        "incremental_cert_path": incremental_cert_path,
        "incremental_mapping_path": incremental_mapping_path,
        "review_todo_path": review_path,
    }


def main():
    config_path = os.path.join(PROJECT_ROOT, "task_customer", "config", "config.yaml")
    for arg in sys.argv[1:]:
        if arg.startswith("--config="):
            config_path = arg.split("=", 1)[1].strip()
    run_incremental(config_path)


if __name__ == "__main__":
    main()
