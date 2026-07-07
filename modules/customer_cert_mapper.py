import os
import re
import yaml
import pandas as pd
from difflib import SequenceMatcher
from datetime import datetime
from time import perf_counter


class CustomerCertMapper:
    def __init__(self, config_path: str):
        with open(config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f) or {}
        self.paths = self.config.get("paths", {}) or {}
        self.columns = self.config.get("columns", {}) or {}
        self.analysis_cfg = self.config.get("analysis", {}) or {}

    def _normalize(self, value):
        return "".join(str(value).replace("\u3000", " ").split()).strip()

    def _pick_col(self, df, candidates, required=True, label=""):
        if df is None or df.empty:
            if required:
                raise ValueError(f"数据为空，无法识别字段：{label}")
            return None
        norm_map = {self._normalize(c): c for c in df.columns}
        for cand in candidates or []:
            key = self._normalize(cand)
            if key in norm_map:
                return norm_map[key]
        if required:
            cols_preview = ", ".join([str(c) for c in list(df.columns)[:40]])
            raise ValueError(f"缺少必要字段：{label}。候选={candidates}。实际列={cols_preview}")
        return None

    def _read_excel(self, path, label):
        if not path:
            raise ValueError(f"缺少文件路径配置：{label}")
        if not os.path.exists(path):
            raise FileNotFoundError(f"找不到文件：{path}")
        return pd.read_excel(path)

    def _log(self, msg):
        if not self.analysis_cfg.get("cert_mapping_enable_timing_log", True):
            return
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{ts}] {msg}")

    def _clean_company_name(self, name):
        s = str(name or "").strip().lower()
        if not s or s == "nan":
            return ""
        s = re.sub(r"[\(\)\[\]（）【】]", " ", s)
        s = re.sub(r"[^a-z0-9\u4e00-\u9fff\s\-&/]", " ", s)
        tokens = [t for t in re.split(r"[\s\-/&]+", s) if t]
        noise = set([str(x).strip().lower() for x in (self.analysis_cfg.get("cert_mapping_suffix_noise", []) or []) if str(x).strip()])
        tokens = [t for t in tokens if t not in noise]
        return " ".join(tokens).strip()

    def _clean_org_unit(self, value):
        s = str(value or "").strip().lower()
        if not s or s == "nan":
            return ""
        s = s.replace("&", " and ")
        s = re.sub(r"[^a-z0-9\u4e00-\u9fff\s]", " ", s)
        tokens = [t for t in s.split() if t]
        noise = {
            "group", "region", "department", "office", "rep", "representative",
            "代表处", "大区", "区域", "部门"
        }
        tokens = [t for t in tokens if t not in noise]
        return " ".join(tokens).strip()

    def _same_text(self, left, right):
        left_val = str(left or "").strip()
        right_val = str(right or "").strip()
        if not left_val or not right_val:
            return True
        return left_val == right_val

    def _same_org_unit(self, left, right):
        left_val = self._clean_org_unit(left)
        right_val = self._clean_org_unit(right)
        if not left_val or not right_val:
            return True
        return left_val == right_val

    def _filter_records_by_org(self, records, region_key, sub_region_key=None, country_key=None,
                               region_value="", sub_region_value="", country_value=""):
        filtered = []
        for row in records:
            if region_value and not self._same_text(row.get(region_key, ""), region_value):
                continue
            if sub_region_key and sub_region_value and not self._same_org_unit(row.get(sub_region_key, ""), sub_region_value):
                continue
            if country_key and country_value and not self._same_org_unit(row.get(country_key, ""), country_value):
                continue
            filtered.append(row)
        return filtered

    def _token_similarity(self, a, b):
        if not a or not b:
            return 0.0
        a_set = set(a.split())
        b_set = set(b.split())
        inter = len(a_set & b_set)
        union = len(a_set | b_set)
        jaccard = inter / union if union else 0.0
        seq = SequenceMatcher(None, a, b).ratio()
        return round(max(jaccard, seq), 4)

    def _score_match(self, source_clean, target_clean):
        if not source_clean or not target_clean:
            return 0.0, "unmatched"
        if source_clean == target_clean:
            return 1.0, "normalized_exact"
        if source_clean in target_clean or target_clean in source_clean:
            return 0.85, "normalized_contains"
        score = self._token_similarity(source_clean, target_clean)
        return score, "token_similarity"

    def _tokenize(self, value):
        s = str(value or "").strip()
        return [t for t in s.split() if t]

    def _build_index(self, records, clean_key):
        by_exact = {}
        token_index = {}
        for row in records:
            clean = str(row.get(clean_key, "")).strip()
            if not clean:
                continue
            by_exact.setdefault(clean, []).append(row)
            for token in set(self._tokenize(clean)):
                token_index.setdefault(token, []).append(row)
        return by_exact, token_index

    def _lookup_candidates(self, name_clean, by_exact, token_index, max_candidates=200):
        if not name_clean:
            return []
        exact = by_exact.get(name_clean, [])
        if exact:
            return exact
        scored = {}
        for token in set(self._tokenize(name_clean)):
            for row in token_index.get(token, []):
                key = id(row)
                scored[key] = row
        candidates = list(scored.values())
        if len(candidates) <= max_candidates:
            return candidates
        candidates = sorted(
            candidates,
            key=lambda r: self._token_similarity(name_clean, str(r.get("target_name_clean", r.get("source_name_clean", "")))),
            reverse=True,
        )
        return candidates[:max_candidates]

    def _build_target_customers(self, customer_df):
        cust_cols = self.columns.get("customer", {}) or {}
        cust_id_col = self._pick_col(customer_df, cust_cols.get("customer_id_candidates"), label="目标客户ID")
        cust_name_col = self._pick_col(customer_df, cust_cols.get("customer_name_candidates"), label="目标客户名称")
        region_col = self._pick_col(customer_df, cust_cols.get("region_l2_candidates"), label="目标客户大区")
        sub_region_col = self._pick_col(customer_df, cust_cols.get("region_l3_candidates"), required=False, label="目标客户代表处")
        account_type_col = self._pick_col(customer_df, cust_cols.get("account_type_candidates"), required=False, label="目标客户类型")

        target = pd.DataFrame()
        target["target_customer_id"] = customer_df[cust_id_col].astype(str).str.strip()
        target["target_customer_name"] = customer_df[cust_name_col].astype(str).str.strip()
        target["target_region"] = customer_df[region_col].astype(str).str.strip()
        target["target_sub_region"] = customer_df[sub_region_col].astype(str).str.strip() if sub_region_col else ""
        target["account_type"] = customer_df[account_type_col].astype(str).str.strip() if account_type_col else ""
        target = target[target["target_customer_id"].str.len() > 0].copy()
        account_type_filter = str(self.analysis_cfg.get("account_type_filter", "") or "").strip()
        if account_type_filter and "account_type" in target.columns:
            target = target[target["account_type"] == account_type_filter].copy()
        ignore_regions = set([str(x).strip() for x in (self.analysis_cfg.get("ignore_regions", []) or []) if str(x).strip()])
        if ignore_regions:
            target = target[~target["target_region"].isin(ignore_regions)].copy()
        region_name_map = self.analysis_cfg.get("region_name_map", {}) or {}
        if region_name_map:
            target["target_region"] = target["target_region"].apply(lambda x: region_name_map.get(str(x).strip(), str(x).strip()))
        target["target_name_clean"] = target["target_customer_name"].apply(self._clean_company_name)
        target["target_sub_region_clean"] = target["target_sub_region"].apply(self._clean_org_unit)
        target = target.drop_duplicates(subset=["target_customer_id"])
        return target

    def _build_source_customers(self, full_customer_df):
        full_cols = self.columns.get("full_customer", {}) or {}
        src_id_col = self._pick_col(full_customer_df, full_cols.get("customer_id_candidates"), required=False, label="全量客户ID")
        src_account_col = self._pick_col(full_customer_df, full_cols.get("account_code_candidates"), required=False, label="全量客户账号")
        src_name_col = self._pick_col(full_customer_df, full_cols.get("customer_name_candidates"), label="全量客户名称")
        src_region_col = self._pick_col(full_customer_df, full_cols.get("region_l2_candidates"), required=False, label="全量客户大区")
        src_sub_region_col = self._pick_col(full_customer_df, full_cols.get("region_l3_candidates"), required=False, label="全量客户代表处")
        src_country_col = self._pick_col(full_customer_df, full_cols.get("country_candidates"), required=False, label="全量客户国家")

        source = pd.DataFrame()
        source["source_customer_id"] = full_customer_df[src_id_col].astype(str).str.strip() if src_id_col else ""
        source["source_account_code"] = full_customer_df[src_account_col].astype(str).str.strip() if src_account_col else ""
        source["source_customer_name"] = full_customer_df[src_name_col].astype(str).str.strip()
        source["source_region_cn"] = full_customer_df[src_region_col].astype(str).str.strip() if src_region_col else ""
        source["source_sub_region_cn"] = full_customer_df[src_sub_region_col].astype(str).str.strip() if src_sub_region_col else ""
        source["source_country"] = full_customer_df[src_country_col].astype(str).str.strip() if src_country_col else ""
        source = source[source["source_customer_name"].str.len() > 0].copy()

        cn_region_map = self.analysis_cfg.get("cert_mapping_cn_region_map", {}) or {}
        source["source_region"] = source["source_region_cn"].apply(lambda x: cn_region_map.get(str(x).strip(), str(x).strip()))
        source["source_sub_region_clean"] = source["source_sub_region_cn"].apply(self._clean_org_unit)
        source["source_country_clean"] = source["source_country"].apply(self._clean_org_unit)
        source["source_name_clean"] = source["source_customer_name"].apply(self._clean_company_name)
        source = source.drop_duplicates(subset=["source_customer_id", "source_account_code", "source_customer_name"])
        return source

    def _build_cert(self, cert_df):
        cert_cols = self.columns.get("cert", {}) or {}
        cert_id_col = self._pick_col(cert_df, cert_cols.get("customer_id_candidates"), label="证书客户ID")
        cert_name_col = self._pick_col(cert_df, cert_cols.get("customer_name_candidates"), required=False, label="证书客户名称")
        cert_type_col = self._pick_col(cert_df, cert_cols.get("cert_type_candidates"), required=False, label="证书类型")
        cert_region_col = self._pick_col(cert_df, cert_cols.get("region_l2_candidates"), required=False, label="证书客户大区")
        cert_sub_region_col = self._pick_col(cert_df, cert_cols.get("region_l3_candidates"), required=False, label="证书客户代表处")
        cert_country_col = self._pick_col(cert_df, cert_cols.get("country_candidates"), required=False, label="证书客户国家")

        cert = pd.DataFrame()
        cert["cert_gsp_id"] = cert_df[cert_id_col].astype(str).str.strip()
        cert["cert_company_name"] = cert_df[cert_name_col].astype(str).str.strip() if cert_name_col else ""
        cert["cert_type"] = cert_df[cert_type_col].astype(str).str.strip() if cert_type_col else ""
        cert["cert_region_raw"] = cert_df[cert_region_col].astype(str).str.strip() if cert_region_col else ""
        cert["cert_sub_region_raw"] = cert_df[cert_sub_region_col].astype(str).str.strip() if cert_sub_region_col else ""
        cert["cert_country_raw"] = cert_df[cert_country_col].astype(str).str.strip() if cert_country_col else ""
        cert = cert[cert["cert_gsp_id"].str.len() > 0].copy()
        region_name_map = self.analysis_cfg.get("region_name_map", {}) or {}
        cert["cert_region"] = cert["cert_region_raw"].apply(lambda x: region_name_map.get(str(x).strip(), str(x).strip()))
        cert["cert_sub_region_clean"] = cert["cert_sub_region_raw"].apply(self._clean_org_unit)
        cert["cert_country_clean"] = cert["cert_country_raw"].apply(self._clean_org_unit)
        cert["cert_name_clean"] = cert["cert_company_name"].apply(self._clean_company_name)
        cert = cert.drop_duplicates(subset=["cert_gsp_id", "cert_company_name", "cert_type"])
        return cert

    def build_mapping(self):
        t0 = perf_counter()
        self._log("证书映射：开始")
        customer_df = self._read_excel(self.paths.get("customer_file"), "customer_file")
        cert_df = self._read_excel(self.paths.get("cert_file"), "cert_file")
        full_path = self.paths.get("full_customer_file")
        if full_path and os.path.exists(full_path):
            full_customer_df = self._read_excel(full_path, "full_customer_file")
        else:
            full_customer_df = customer_df.copy()
        self._log(f"证书映射：读取数据完成，用时 {perf_counter() - t0:.2f}s")

        t1 = perf_counter()
        targets = self._build_target_customers(customer_df)
        sources = self._build_source_customers(full_customer_df)
        cert = self._build_cert(cert_df)
        self._log(f"证书映射：清洗与预处理完成，用时 {perf_counter() - t1:.2f}s")

        same_region_only = bool(self.analysis_cfg.get("cert_mapping_same_region_only", True))
        threshold = float(self.analysis_cfg.get("cert_mapping_similarity_threshold", 0.45))
        low_conf_threshold = float(self.analysis_cfg.get("cert_mapping_low_confidence_threshold", 0.65))
        basis_labels = self.analysis_cfg.get("cert_mapping_basis_labels", {}) or {}
        progress_every = int(self.analysis_cfg.get("cert_mapping_progress_every", 50) or 50)

        t2 = perf_counter()
        target_id_set = set(targets["target_customer_id"].tolist())
        target_records = targets.to_dict("records")
        targets_by_region = {}
        for row in target_records:
            targets_by_region.setdefault(str(row.get("target_region", "")).strip(), []).append(row)
        target_exact_global, target_token_global = self._build_index(target_records, "target_name_clean")
        target_index_by_region = {}
        for region, rows in targets_by_region.items():
            target_index_by_region[region] = self._build_index(rows, "target_name_clean")
        self._log(f"证书映射：建立目标客户索引完成，用时 {perf_counter() - t2:.2f}s")

        t3 = perf_counter()
        source_by_account = {}
        source_records = sources.to_dict("records")
        for row in source_records:
            key = str(row.get("source_account_code", "")).strip()
            if key:
                source_by_account.setdefault(key, []).append(row)
            sid = str(row.get("source_customer_id", "")).strip()
            if sid:
                source_by_account.setdefault(sid, []).append(row)
        source_exact, source_token_index = self._build_index(source_records, "source_name_clean")
        self._log(f"证书映射：建立全量客户索引完成，用时 {perf_counter() - t3:.2f}s")

        mapping_rows = []
        low_conf_rows = []
        unmatched_rows = []

        t4 = perf_counter()
        cert_records = cert.to_dict("records")
        total = len(cert_records)
        self._log(f"证书映射：开始匹配，证书记录数={total}")

        for i, cert_row in enumerate(cert_records, start=1):
            if progress_every > 0 and (i == 1 or i % progress_every == 0 or i == total):
                elapsed = perf_counter() - t4
                rate = (i / elapsed) if elapsed > 0 else 0.0
                self._log(f"证书映射：进度 {i}/{total}，已用 {elapsed:.1f}s，{rate:.2f} 条/秒")

            gsp_id = str(cert_row.get("cert_gsp_id", "")).strip()
            cert_name = str(cert_row.get("cert_company_name", "")).strip()
            cert_name_clean = str(cert_row.get("cert_name_clean", "")).strip()
            cert_region = str(cert_row.get("cert_region", "")).strip()
            cert_sub_region_raw = str(cert_row.get("cert_sub_region_raw", "")).strip()
            cert_sub_region_clean = str(cert_row.get("cert_sub_region_clean", "")).strip()
            cert_country_raw = str(cert_row.get("cert_country_raw", "")).strip()
            cert_country_clean = str(cert_row.get("cert_country_clean", "")).strip()

            if gsp_id in target_id_set:
                target_row = targets[targets["target_customer_id"] == gsp_id].iloc[0]
                target_region = str(target_row["target_region"]).strip()
                target_sub_region = str(target_row["target_sub_region"]).strip()
                if cert_region and target_region and cert_region != target_region:
                    unmatched_rows.append({
                        "cert_gsp_id": gsp_id,
                        "cert_company_name": cert_name,
                        "cert_region": cert_region,
                        "cert_sub_region": cert_sub_region_raw,
                        "source_customer_name": cert_name,
                        "source_region": target_region,
                        "best_score": 1.0,
                        "reason": "证书客户ID直接命中目标客户，但大区不一致，已阻止自动匹配",
                    })
                    continue
                if cert_sub_region_clean and self._clean_org_unit(target_sub_region) and cert_sub_region_clean != self._clean_org_unit(target_sub_region):
                    unmatched_rows.append({
                        "cert_gsp_id": gsp_id,
                        "cert_company_name": cert_name,
                        "cert_region": cert_region,
                        "cert_sub_region": cert_sub_region_raw,
                        "source_customer_name": cert_name,
                        "source_region": target_region,
                        "best_score": 1.0,
                        "reason": "证书客户ID直接命中目标客户，但代表处不一致，已阻止自动匹配",
                    })
                    continue
                mapping_rows.append({
                    "cert_gsp_id": gsp_id,
                    "cert_company_name": cert_name,
                    "cert_region": cert_region,
                    "cert_sub_region": cert_sub_region_raw,
                    "source_customer_id": gsp_id,
                    "source_customer_name": cert_name or str(target_row["target_customer_name"]),
                    "source_region": target_region,
                    "source_sub_region": target_sub_region,
                    "mapped_customer_id": str(target_row["target_customer_id"]),
                    "mapped_customer_name": str(target_row["target_customer_name"]),
                    "mapped_region": target_region,
                    "mapped_sub_region": target_sub_region,
                    "match_score": 1.0,
                    "match_basis": basis_labels.get("direct_id_match", "账号直接匹配"),
                    "match_status": "matched",
                })
                continue

            source_candidates = source_by_account.get(gsp_id, [])
            if not source_candidates and cert_name_clean:
                source_candidates = self._lookup_candidates(cert_name_clean, source_exact, source_token_index, max_candidates=50)
            source_candidates = self._filter_records_by_org(
                source_candidates,
                region_key="source_region",
                sub_region_key="source_sub_region_cn",
                country_key="source_country",
                region_value=cert_region,
                sub_region_value=cert_sub_region_raw,
                country_value=cert_country_raw,
            )

            best_source_row = None
            best_target_row = None
            best_score = 0.0
            best_basis = "unmatched"
            ranked_sources = sorted(
                source_candidates,
                key=lambda r: self._token_similarity(cert_name_clean, str(r.get("source_name_clean", ""))),
                reverse=True,
            ) if source_candidates else []

            for src in ranked_sources or [{}]:
                source_name_clean = str(src.get("source_name_clean", "")).strip() or cert_name_clean
                source_region = str(src.get("source_region", "")).strip()
                source_sub_region = str(src.get("source_sub_region_cn", "")).strip()

                effective_region = cert_region or source_region
                target_candidates = []
                if same_region_only and effective_region:
                    idx = target_index_by_region.get(effective_region)
                    if idx:
                        target_candidates.extend(self._lookup_candidates(source_name_clean, idx[0], idx[1], max_candidates=80))
                elif not cert_region:
                    target_candidates = self._lookup_candidates(source_name_clean, target_exact_global, target_token_global, max_candidates=120)

                if not target_candidates and same_region_only and effective_region:
                    target_candidates.extend(targets_by_region.get(effective_region, []))
                if not target_candidates and not effective_region:
                    target_candidates = target_records

                target_candidates = self._filter_records_by_org(
                    target_candidates,
                    region_key="target_region",
                    sub_region_key="target_sub_region",
                    region_value=effective_region,
                    sub_region_value=cert_sub_region_raw or source_sub_region,
                )

                if src:
                    target_candidates = self._filter_records_by_org(
                        target_candidates,
                        region_key="target_region",
                        sub_region_key="target_sub_region",
                        region_value=source_region,
                        sub_region_value=source_sub_region,
                    )

                dedup_targets = {}
                for tgt in target_candidates:
                    dedup_targets[str(tgt.get("target_customer_id", ""))] = tgt
                for tgt in dedup_targets.values():
                    score, basis = self._score_match(source_name_clean, str(tgt.get("target_name_clean", "")))
                    if score > best_score:
                        best_score = score
                        best_basis = basis
                        best_source_row = src if src else {}
                        best_target_row = tgt

            if best_target_row is None or best_score < threshold:
                unmatched_rows.append({
                    "cert_gsp_id": gsp_id,
                    "cert_company_name": cert_name,
                    "cert_region": cert_region,
                    "cert_sub_region": cert_sub_region_raw,
                    "source_customer_name": ranked_sources[0].get("source_customer_name", "") if ranked_sources else "",
                    "source_region": ranked_sources[0].get("source_region", "") if ranked_sources else "",
                    "best_score": best_score,
                    "reason": "未达到匹配阈值、无候选或组织信息约束不一致",
                })
                continue

            basis_key = best_basis
            basis_desc = basis_labels.get(basis_key, basis_key)
            source_ref = best_source_row if best_source_row else {}
            row = {
                "cert_gsp_id": gsp_id,
                "cert_company_name": cert_name,
                "cert_region": cert_region,
                "cert_sub_region": cert_sub_region_raw,
                "source_customer_id": str(source_ref.get("source_customer_id", "")).strip(),
                "source_customer_name": str(source_ref.get("source_customer_name", cert_name)).strip(),
                "source_region": str(source_ref.get("source_region", "")).strip(),
                "source_sub_region": str(source_ref.get("source_sub_region_cn", "")).strip(),
                "mapped_customer_id": str(best_target_row.get("target_customer_id", "")),
                "mapped_customer_name": str(best_target_row.get("target_customer_name", "")),
                "mapped_region": str(best_target_row.get("target_region", "")),
                "mapped_sub_region": str(best_target_row.get("target_sub_region", "")),
                "match_score": round(float(best_score), 4),
                "match_basis": basis_desc,
                "match_status": "matched",
            }
            mapping_rows.append(row)
            if best_score < low_conf_threshold:
                low_conf_rows.append(row.copy())

        self._log(f"证书映射：匹配完成，用时 {perf_counter() - t4:.2f}s")

        mapping_df = pd.DataFrame(mapping_rows)
        low_conf_df = pd.DataFrame(low_conf_rows)
        unmatched_df = pd.DataFrame(unmatched_rows)

        if not mapping_df.empty:
            mapping_df = mapping_df.drop_duplicates(subset=["cert_gsp_id", "mapped_customer_id"])
        if not low_conf_df.empty:
            low_conf_df = low_conf_df.drop_duplicates(subset=["cert_gsp_id", "mapped_customer_id"])
        if not unmatched_df.empty:
            unmatched_df = unmatched_df.drop_duplicates(subset=["cert_gsp_id"])

        self._log(f"证书映射：结束，总用时 {perf_counter() - t0:.2f}s")
        return {
            "mapping_df": mapping_df,
            "low_confidence_df": low_conf_df,
            "unmatched_df": unmatched_df,
            "targets_df": targets,
            "sources_df": sources,
            "cert_df": cert,
        }
