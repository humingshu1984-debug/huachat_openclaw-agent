import os
import time
import yaml
import pandas as pd
from html.parser import HTMLParser


class CustomerCoverageAnalyzer:
    def __init__(self, config_path: str):
        self.config_path = config_path
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

    def _split_courses(self, value):
        if value is None:
            return []
        s = str(value).strip()
        if not s or s.lower() == "nan":
            return []
        parts = [p.strip() for p in s.split(";")]
        return [p for p in parts if p]

    def _load_cert_mapping(self):
        mapping_path = self.paths.get("cert_mapping_file")
        if not mapping_path or not os.path.exists(mapping_path):
            return pd.DataFrame()
        try:
            return pd.read_excel(mapping_path, sheet_name="mapping")
        except Exception:
            return pd.read_excel(mapping_path)

    def _clean_name_key(self, value):
        s = str(value or "").strip().lower()
        if not s or s == "nan":
            return ""
        keep = []
        for ch in s:
            if ch.isalnum() or ("\u4e00" <= ch <= "\u9fff"):
                keep.append(ch)
            else:
                keep.append(" ")
        return " ".join("".join(keep).split()).strip()

    def _read_msg_html(self, path):
        if not path or not os.path.exists(path):
            return ""
        try:
            import win32com.client as win32

            outlook = win32.Dispatch("Outlook.Application")
            session = outlook.Session
            abs_path = os.path.abspath(path)
            item = None
            try:
                item = session.OpenSharedItem(abs_path)
            except Exception:
                item = None

            if item is None:
                try:
                    data = open(abs_path, "rb").read()
                    tmp_dir = os.path.dirname(abs_path)
                    tmp_name = f"_latin_tmp_{time.time_ns()}.msg"
                    tmp_path = os.path.join(tmp_dir, tmp_name)
                    open(tmp_path, "wb").write(data)
                    try:
                        item = session.OpenSharedItem(os.path.abspath(tmp_path))
                        html = getattr(item, "HTMLBody", "") or ""
                        try:
                            item.Close(0)
                        except Exception:
                            pass
                        try:
                            os.remove(tmp_path)
                        except Exception:
                            pass
                        return html
                    except Exception:
                        try:
                            os.remove(tmp_path)
                        except Exception:
                            pass
                except Exception:
                    item = None

            if item is not None:
                html = getattr(item, "HTMLBody", "") or ""
                try:
                    item.Close(0)
                except Exception:
                    pass
                return html
        except Exception:
            pass

        try:
            data = open(path, "rb").read()
        except Exception:
            return ""

        chunks = []
        lower = data.lower()
        start = 0
        while True:
            idx = lower.find(b"<html", start)
            if idx < 0:
                break
            end = lower.find(b"</html>", idx)
            if end < 0:
                break
            chunks.append(data[idx : end + 7])
            start = end + 7
        if not chunks:
            return ""

        best = max(chunks, key=lambda b: b.count(b"<table"))
        for enc in ("utf-8", "utf-16-le", "latin1"):
            try:
                txt = best.decode(enc, errors="ignore")
                if "<table" in txt.lower():
                    return txt
            except Exception:
                continue
        return ""

    def _extract_table_texts(self, html):
        if not html:
            return []

        class _TableTextParser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.in_cell = False
                self.texts = []
                self._buf = []

            def handle_starttag(self, tag, attrs):
                t = tag.lower()
                if t in ("td", "th"):
                    self.in_cell = True
                    self._buf = []

            def handle_endtag(self, tag):
                t = tag.lower()
                if t in ("td", "th") and self.in_cell:
                    s = "".join(self._buf).strip()
                    s = " ".join(s.split())
                    if s:
                        self.texts.append(s)
                    self._buf = []
                    self.in_cell = False

            def handle_data(self, data):
                if self.in_cell and data:
                    self._buf.append(str(data))

        p = _TableTextParser()
        try:
            p.feed(html)
        except Exception:
            return []
        return p.texts

    def _latin_np_extra_covered_ids(self, customer_dim):
        latin = customer_dim[customer_dim["region_l2"].astype(str).str.strip() == "拉美区"].copy()
        if latin.empty:
            return set()

        required = bool(self.analysis_cfg.get("latin_np_coverage_msg_required", False))
        msg_path = self.paths.get("latin_np_coverage_msg")
        if not msg_path or not os.path.exists(msg_path):
            if required:
                raise FileNotFoundError(f"拉美区覆盖统计需要 .msg 文件，但找不到文件：paths.latin_np_coverage_msg={msg_path}")
            return set()

        html = self._read_msg_html(msg_path)
        if not html:
            if required:
                raise ValueError(f"拉美区 .msg 邮件正文解析失败或为空：{msg_path}")
            return set()

        cells = self._extract_table_texts(html)
        if not cells:
            if required:
                raise ValueError(f"拉美区 .msg 邮件正文未包含可解析的表格单元格：{msg_path}")
            return set()

        cell_keys = [self._clean_name_key(c) for c in cells]
        cell_keys = [c for c in cell_keys if c]
        if not cell_keys:
            if required:
                raise ValueError(f"拉美区 .msg 表格文本清洗后为空，无法用于客户匹配：{msg_path}")
            return set()

        matched = set()
        for _, row in latin.iterrows():
            cid = str(row.get("customer_id", "")).strip()
            name = self._clean_name_key(row.get("customer_name", ""))
            if not cid or not name:
                continue
            for ck in cell_keys:
                if name in ck or ck in name:
                    matched.add(cid)
                    break
        return matched

    def analyze(self, use_cert_mapping=True, region_filter=None):
        customer_df = self._read_excel(self.paths.get("customer_file"), "customer_file")
        full_customer_df = None
        full_customer_path = self.paths.get("full_customer_file")
        if full_customer_path and os.path.exists(full_customer_path):
            try:
                full_customer_df = self._read_excel(full_customer_path, "full_customer_file")
            except Exception:
                full_customer_df = None
        activity_df = self._read_excel(self.paths.get("activity_file"), "activity_file")
        activity_course_df = self._read_excel(self.paths.get("activity_course_file"), "activity_course_file")
        focus_course_df = self._read_excel(self.paths.get("focus_course_file"), "focus_course_file")
        cert_df = self._read_excel(self.paths.get("cert_file"), "cert_file")

        customer_cols = self.columns.get("customer", {}) or {}
        full_customer_cols = self.columns.get("full_customer", {}) or {}
        activity_cols = self.columns.get("activity", {}) or {}
        activity_course_cols = self.columns.get("activity_course", {}) or {}
        focus_course_cols = self.columns.get("focus_course", {}) or {}
        cert_cols = self.columns.get("cert", {}) or {}

        cust_id_col = self._pick_col(customer_df, customer_cols.get("customer_id_candidates"), label="客户ID(客户清单)")
        cust_name_col = self._pick_col(customer_df, customer_cols.get("customer_name_candidates"), required=False, label="客户名称(客户清单)")
        region_l2_col = self._pick_col(customer_df, customer_cols.get("region_l2_candidates"), label="二级部门/大区(客户清单)")
        region_l3_col = self._pick_col(customer_df, customer_cols.get("region_l3_candidates"), required=False, label="三级部门/代表处(客户清单)")
        account_type_col = self._pick_col(customer_df, customer_cols.get("account_type_candidates"), required=False, label="客户类型(客户清单)")
        country_col = None
        if customer_cols.get("country_candidates"):
            country_col = self._pick_col(customer_df, customer_cols.get("country_candidates"), required=False, label="国家(客户清单)")

        act_cust_id_col = self._pick_col(activity_df, activity_cols.get("customer_id_candidates"), label="客户ID(活动记录)")
        act_id_col = self._pick_col(activity_df, activity_cols.get("activity_id_candidates"), label="活动ID(活动记录)")
        act_course_col = self._pick_col(activity_df, activity_cols.get("training_courses_candidates"), required=False, label="活动课程(活动记录)")

        detail_act_id_col = self._pick_col(activity_course_df, activity_course_cols.get("activity_id_candidates"), label="活动ID(活动课程明细)")
        detail_course_col = self._pick_col(activity_course_df, activity_course_cols.get("course_name_candidates"), label="课程名称(活动课程明细)")

        focus_course_col = self._pick_col(focus_course_df, focus_course_cols.get("course_name_candidates"), label="课程名称(关注课程列表)")

        cert_cust_id_col = self._pick_col(cert_df, cert_cols.get("customer_id_candidates"), label="客户ID(证书清单)")
        cert_name_col = self._pick_col(cert_df, cert_cols.get("customer_name_candidates"), required=False, label="客户名称(证书清单)")
        cert_type_col = self._pick_col(cert_df, cert_cols.get("cert_type_candidates"), required=False, label="证书类型(证书清单)")

        customer_dim = pd.DataFrame()
        customer_dim["customer_id"] = customer_df[cust_id_col].astype(str).str.strip()
        customer_dim["customer_name"] = customer_df[cust_name_col].astype(str).str.strip() if cust_name_col else ""
        customer_dim["region_l2"] = customer_df[region_l2_col].astype(str).str.strip()
        customer_dim["region_l3"] = customer_df[region_l3_col].astype(str).str.strip() if region_l3_col else "未知"
        customer_dim["account_type"] = customer_df[account_type_col].astype(str).str.strip() if account_type_col else ""
        customer_dim["country"] = customer_df[country_col].astype(str).str.strip() if country_col else ""
        customer_dim["region_l3_cn"] = ""
        customer_dim["country_cn"] = ""
        customer_dim = customer_dim[customer_dim["customer_id"].str.len() > 0].copy()
        customer_dim.loc[customer_dim["region_l2"].isin(["", "nan", "None"]), "region_l2"] = "未知"
        customer_dim.loc[customer_dim["region_l3"].isin(["", "nan", "None"]), "region_l3"] = "未知"

        if full_customer_df is not None and not full_customer_df.empty:
            try:
                full_account_code_col = self._pick_col(
                    full_customer_df,
                    full_customer_cols.get("account_code_candidates"),
                    required=False,
                    label="Account Code(全量客户清单)",
                )
                full_region_l3_cn_col = self._pick_col(
                    full_customer_df,
                    full_customer_cols.get("region_l3_candidates"),
                    required=False,
                    label="代表处CN(全量客户清单)",
                )
                full_country_col = self._pick_col(
                    full_customer_df,
                    full_customer_cols.get("country_candidates"),
                    required=False,
                    label="国家(全量客户清单)",
                )
                if full_account_code_col:
                    full_map = pd.DataFrame()
                    full_map["customer_id"] = full_customer_df[full_account_code_col].astype(str).str.strip()
                    full_map["region_l3_cn_full"] = full_customer_df[full_region_l3_cn_col].astype(str).str.strip() if full_region_l3_cn_col else ""
                    full_map["country_cn_full"] = full_customer_df[full_country_col].astype(str).str.strip() if full_country_col else ""
                    full_map = full_map[full_map["customer_id"].str.len() > 0].drop_duplicates(subset=["customer_id"])
                    customer_dim = customer_dim.merge(full_map, on="customer_id", how="left")
                    customer_dim["region_l3_cn"] = customer_dim["region_l3_cn"].fillna("").astype(str).str.strip()
                    customer_dim["country_cn"] = customer_dim["country_cn"].fillna("").astype(str).str.strip()
                    if "region_l3_cn_full" in customer_dim.columns:
                        src = customer_dim["region_l3_cn_full"].fillna("").astype(str).str.strip()
                        customer_dim.loc[customer_dim["region_l3_cn"].isin(["", "nan", "None"]), "region_l3_cn"] = src
                        customer_dim = customer_dim.drop(columns=["region_l3_cn_full"])
                    if "country_cn_full" in customer_dim.columns:
                        src = customer_dim["country_cn_full"].fillna("").astype(str).str.strip()
                        customer_dim.loc[customer_dim["country_cn"].isin(["", "nan", "None"]), "country_cn"] = src
                        customer_dim = customer_dim.drop(columns=["country_cn_full"])
            except Exception:
                pass

        account_type_filter = str(self.analysis_cfg.get("account_type_filter", "") or "").strip()
        if account_type_filter and "account_type" in customer_dim.columns:
            customer_dim = customer_dim[customer_dim["account_type"] == account_type_filter].copy()

        ignore_regions = set([str(x).strip() for x in (self.analysis_cfg.get("ignore_regions", []) or []) if str(x).strip()])
        if ignore_regions:
            customer_dim = customer_dim[~customer_dim["region_l2"].isin(ignore_regions)].copy()

        region_name_map = self.analysis_cfg.get("region_name_map", {}) or {}
        if region_name_map:
            customer_dim["region_l2"] = customer_dim["region_l2"].apply(lambda x: region_name_map.get(str(x).strip(), str(x).strip()))

        if region_filter:
            region_set = set([str(x).strip() for x in region_filter if str(x).strip()])
            if region_set:
                customer_dim = customer_dim[customer_dim["region_l2"].isin(region_set)].copy()

        customer_dim = customer_dim.drop_duplicates(subset=["customer_id"])

        focus_courses = set(
            focus_course_df[focus_course_col]
            .astype(str)
            .str.strip()
            .replace("nan", "")
            .tolist()
        )
        focus_courses = {c for c in focus_courses if c}

        activity = pd.DataFrame()
        activity["customer_id"] = activity_df[act_cust_id_col].astype(str).str.strip()
        activity["activity_id"] = activity_df[act_id_col].astype(str).str.strip()
        if act_course_col:
            activity["training_courses_raw"] = activity_df[act_course_col].astype(str).str.strip()
        else:
            activity["training_courses_raw"] = ""
        activity = activity[activity["customer_id"].str.len() > 0].copy()

        activity = activity.merge(
            customer_dim[["customer_id", "customer_name", "region_l2", "region_l3", "account_type", "country", "region_l3_cn", "country_cn"]],
            on="customer_id",
            how="inner",
        )

        detail_courses = pd.DataFrame()
        detail_courses["activity_id"] = activity_course_df[detail_act_id_col].astype(str).str.strip()
        detail_courses["training_courses_raw"] = activity_course_df[detail_course_col].astype(str).str.strip()
        detail_courses = detail_courses[detail_courses["activity_id"].str.len() > 0].copy()

        detail_course_rows = []
        for _, row in detail_courses.iterrows():
            for course_name in self._split_courses(row["training_courses_raw"]):
                detail_course_rows.append({"activity_id": row["activity_id"], "course_name": course_name})
        detail_course_items = pd.DataFrame(detail_course_rows) if detail_course_rows else pd.DataFrame(columns=["activity_id", "course_name"])

        if detail_course_items.empty and act_course_col:
            fallback_rows = []
            for _, row in activity[["activity_id", "training_courses_raw"]].iterrows():
                for course_name in self._split_courses(row["training_courses_raw"]):
                    fallback_rows.append({"activity_id": row["activity_id"], "course_name": course_name})
            detail_course_items = pd.DataFrame(fallback_rows) if fallback_rows else pd.DataFrame(columns=["activity_id", "course_name"])

        detail_course_items["is_focus_course"] = detail_course_items["course_name"].apply(lambda x: 1 if str(x).strip() in focus_courses else 0)

        activity_focus = activity[["customer_id", "activity_id"]].drop_duplicates().merge(
            detail_course_items, on="activity_id", how="left"
        )
        activity_focus["is_focus_course"] = activity_focus["is_focus_course"].fillna(0).astype(int)

        focus_activity = activity_focus[activity_focus["is_focus_course"] == 1].copy()
        covered_customers = set(focus_activity["customer_id"].dropna().astype(str).str.strip().unique().tolist())
        latin_msg_covered_ids = self._latin_np_extra_covered_ids(customer_dim)
        if latin_msg_covered_ids:
            covered_customers = covered_customers | latin_msg_covered_ids

        activity_cnt = focus_activity.groupby("customer_id")["activity_id"].nunique().rename("training_activity_count").reset_index()
        course_cnt = focus_activity.groupby("customer_id")["course_name"].nunique().rename("focus_course_count").reset_index()
        focus_course_list = (
            focus_activity.groupby("customer_id")["course_name"]
            .apply(lambda xs: "; ".join(sorted(set([str(x).strip() for x in xs if str(x).strip()]))))
            .rename("focus_courses")
            .reset_index()
        )

        cert = pd.DataFrame()
        cert["raw_customer_id"] = cert_df[cert_cust_id_col].astype(str).str.strip()
        cert["customer_id"] = cert["raw_customer_id"]
        cert["cert_company_name"] = cert_df[cert_name_col].astype(str).str.strip() if cert_name_col else ""
        cert["cert_type"] = cert_df[cert_type_col].astype(str).str.strip() if cert_type_col else "Certificate"
        cert = cert[cert["raw_customer_id"].str.len() > 0].copy()

        cert_mapping_used_df = pd.DataFrame(columns=["cert_gsp_id", "cert_company_name", "mapped_customer_id", "mapped_customer_name", "mapped_region", "mapped_sub_region", "match_score", "match_basis", "match_status"])
        if use_cert_mapping:
            cert_mapping_df = self._load_cert_mapping()
            if cert_mapping_df is not None and not cert_mapping_df.empty:
                cert_mapping_df["cert_gsp_id"] = cert_mapping_df["cert_gsp_id"].astype(str).str.strip()
                cert_mapping_df["mapped_customer_id"] = cert_mapping_df["mapped_customer_id"].astype(str).str.strip()
                cert = cert.merge(
                    cert_mapping_df[["cert_gsp_id", "mapped_customer_id", "mapped_customer_name", "mapped_region", "mapped_sub_region", "match_score", "match_basis", "match_status"]],
                    left_on="raw_customer_id",
                    right_on="cert_gsp_id",
                    how="left",
                )
                cert["customer_id"] = cert["mapped_customer_id"].fillna(cert["raw_customer_id"]).astype(str).str.strip()
                cert_mapping_used_df = cert[cert["mapped_customer_id"].notna()].copy()
                if not cert_mapping_used_df.empty:
                    cert_mapping_used_df = cert_mapping_used_df[["raw_customer_id", "cert_company_name", "mapped_customer_id", "mapped_customer_name", "mapped_region", "mapped_sub_region", "match_score", "match_basis", "match_status"]]
                    cert_mapping_used_df = cert_mapping_used_df.rename(columns={"raw_customer_id": "cert_gsp_id"})

        cert = cert.merge(customer_dim[["customer_id", "region_l2", "region_l3", "customer_name"]], on="customer_id", how="inner")

        cert_flag = set(cert["customer_id"].dropna().astype(str).str.strip().unique().tolist())
        cert_types_by_customer = cert.groupby("customer_id")["cert_type"].apply(
            lambda x: "; ".join(sorted(set([v for v in x if str(v).strip()])))
        ).rename("cert_types").reset_index()
        cert_basis_by_customer = cert.groupby("customer_id")["match_basis"].apply(
            lambda x: "; ".join(sorted(set([str(v).strip() for v in x if str(v).strip()])))
        ).rename("cert_mapping_basis").reset_index() if "match_basis" in cert.columns else pd.DataFrame(columns=["customer_id", "cert_mapping_basis"])

        software_keywords = self.analysis_cfg.get("software_cert_keywords", None)
        if software_keywords is None:
            software_keywords = ["Software"]
        software_keywords = [str(k).strip().lower() for k in (software_keywords or []) if str(k).strip()]
        if not software_keywords:
            software_keywords = ["software"]

        def _is_software_cert(v):
            s = str(v or "").strip().lower()
            if not s or s == "nan":
                return False
            for k in software_keywords:
                if k in s:
                    return True
            return False

        cert["is_software_cert"] = cert["cert_type"].apply(_is_software_cert)
        software_cert_flag = set(
            cert[cert["is_software_cert"] == True]["customer_id"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )
        software_cert_types_by_customer = (
            cert[cert["is_software_cert"] == True]
            .groupby("customer_id")["cert_type"]
            .apply(lambda x: "; ".join(sorted(set([str(v).strip() for v in x if str(v).strip()]))))
            .rename("software_cert_types")
            .reset_index()
        ) if software_cert_flag else pd.DataFrame(columns=["customer_id", "software_cert_types"])

        detail = customer_dim.copy()
        detail["covered_training"] = detail["customer_id"].apply(lambda x: 1 if x in covered_customers else 0)
        detail["covered_via_latin_msg"] = detail["customer_id"].apply(lambda x: 1 if x in latin_msg_covered_ids else 0)
        detail["certified"] = detail["customer_id"].apply(lambda x: 1 if x in cert_flag else 0)
        detail["software_certified"] = detail["customer_id"].apply(lambda x: 1 if x in software_cert_flag else 0)
        mapped_ids = set(cert[cert.get("mapped_customer_id", pd.Series(dtype=str)).notna()]["customer_id"].astype(str).str.strip().tolist()) if "mapped_customer_id" in cert.columns else set()
        detail["certified_via_mapping"] = detail["customer_id"].apply(lambda x: 1 if x in mapped_ids else 0)
        detail = detail.merge(activity_cnt, on="customer_id", how="left")
        detail = detail.merge(course_cnt, on="customer_id", how="left")
        detail = detail.merge(focus_course_list, on="customer_id", how="left")
        detail = detail.merge(cert_types_by_customer, on="customer_id", how="left")
        detail = detail.merge(cert_basis_by_customer, on="customer_id", how="left")
        detail = detail.merge(software_cert_types_by_customer, on="customer_id", how="left")
        detail["training_activity_count"] = detail["training_activity_count"].fillna(0).astype(int)
        detail["focus_course_count"] = detail["focus_course_count"].fillna(0).astype(int)
        detail["focus_courses"] = detail["focus_courses"].fillna("")
        detail["cert_types"] = detail["cert_types"].fillna("")
        detail["cert_mapping_basis"] = detail["cert_mapping_basis"].fillna("")
        detail["software_cert_types"] = detail["software_cert_types"].fillna("")

        def _agg_region(df, key_col):
            grouped = df.groupby(key_col).agg(
                np_customer_count=("customer_id", "nunique"),
                covered_count=("covered_training", "sum"),
                certified_count=("certified", "sum"),
                software_certified_count=("software_certified", "sum"),
            ).reset_index()
            grouped["covered_ratio"] = (grouped["covered_count"] / grouped["np_customer_count"] * 100).round(2)
            grouped["certified_ratio"] = (grouped["certified_count"] / grouped["np_customer_count"] * 100).round(2)
            grouped["software_certified_ratio"] = (grouped["software_certified_count"] / grouped["np_customer_count"] * 100).round(2)
            grouped = grouped.sort_values(by=["covered_ratio", "covered_count", key_col], ascending=[False, False, True]).reset_index(drop=True)
            grouped["rank"] = grouped.index + 1
            return grouped

        region_summary = _agg_region(detail, "region_l2")
        region_level3_summary = _agg_region(detail, "region_l3")
        valid_subregion_detail = detail[
            detail["region_l3"].astype(str).str.strip().isin(["", "nan", "None", "未知"]) == False
        ].copy()
        subregion_summary = (
            valid_subregion_detail.groupby(["region_l2", "region_l3"]).agg(
                np_customer_count=("customer_id", "nunique"),
                covered_count=("covered_training", "sum"),
                certified_count=("certified", "sum"),
                software_certified_count=("software_certified", "sum"),
            ).reset_index()
        )
        subregion_summary["covered_ratio"] = (
            subregion_summary["covered_count"] / subregion_summary["np_customer_count"] * 100
        ).round(2)
        subregion_summary["certified_ratio"] = (
            subregion_summary["certified_count"] / subregion_summary["np_customer_count"] * 100
        ).round(2)
        subregion_summary["software_certified_ratio"] = (
            subregion_summary["software_certified_count"] / subregion_summary["np_customer_count"] * 100
        ).round(2)
        subregion_summary = subregion_summary.sort_values(
            by=["region_l2", "covered_ratio", "covered_count", "region_l3"],
            ascending=[True, False, False, True],
        ).reset_index(drop=True)
        subregion_summary["rank_in_region"] = (
            subregion_summary.groupby("region_l2").cumcount() + 1
        )

        cert_type_counts = (
            cert.drop_duplicates(subset=["customer_id", "cert_type"])
            .groupby(["region_l2", "cert_type"])["customer_id"]
            .nunique()
            .reset_index(name="cert_customer_count")
        )
        if cert_type_counts.empty:
            cert_type_count_df = pd.DataFrame(columns=["大区"])
            cert_type_rate_df = pd.DataFrame(columns=["大区"])
        else:
            cert_type_count_df = cert_type_counts.pivot(index="region_l2", columns="cert_type", values="cert_customer_count").fillna(0).astype(int).reset_index()
            cert_type_count_df = cert_type_count_df.rename(columns={"region_l2": "大区"})
            totals = region_summary[["region_l2", "np_customer_count"]].rename(columns={"np_customer_count": "region_total"})
            cert_type_rate = cert_type_counts.merge(totals, on="region_l2", how="left")
            cert_type_rate["cert_type_rate"] = (cert_type_rate["cert_customer_count"] / cert_type_rate["region_total"] * 100).round(2)
            cert_type_rate_df = cert_type_rate.pivot(index="region_l2", columns="cert_type", values="cert_type_rate").fillna(0).reset_index()
            cert_type_rate_df = cert_type_rate_df.rename(columns={"region_l2": "大区"})

        coverage_x_cert_df = region_summary[["region_l2", "np_customer_count", "covered_count", "certified_count"]].copy()
        both = detail[(detail["covered_training"] == 1) & (detail["certified"] == 1)].groupby("region_l2")["customer_id"].nunique().rename("covered_and_certified")
        coverage_x_cert_df = coverage_x_cert_df.merge(both.reset_index(), on="region_l2", how="left")
        coverage_x_cert_df["covered_and_certified"] = coverage_x_cert_df["covered_and_certified"].fillna(0).astype(int)
        coverage_x_cert_df["covered_only"] = coverage_x_cert_df["covered_count"] - coverage_x_cert_df["covered_and_certified"]
        coverage_x_cert_df["certified_only"] = coverage_x_cert_df["certified_count"] - coverage_x_cert_df["covered_and_certified"]
        coverage_x_cert_df["neither"] = (
            coverage_x_cert_df["np_customer_count"]
            - coverage_x_cert_df["covered_and_certified"]
            - coverage_x_cert_df["covered_only"]
            - coverage_x_cert_df["certified_only"]
        )
        coverage_x_cert_df = coverage_x_cert_df.rename(columns={"region_l2": "大区", "np_customer_count": "NP客户数量"})

        top_n = int(self.analysis_cfg.get("top_n", 10))
        top_uncovered_customers_df = detail[detail["covered_training"] == 0].copy()
        top_uncovered_customers_df = top_uncovered_customers_df.sort_values(by=["region_l2", "customer_name"]).head(top_n)

        ppt_region_summary_df = region_summary.rename(
            columns={
                "region_l2": "大区",
                "np_customer_count": "NP客户数量",
                "covered_count": "已覆盖个数",
                "certified_count": "DHIA认证通过个数",
                "software_certified_count": "Software认证通过个数",
                "covered_ratio": "已覆盖比例",
                "certified_ratio": "DHIA认证通过比列",
                "software_certified_ratio": "Software认证通过比例",
            }
        )
        ppt_region_summary_df["已覆盖比例"] = ppt_region_summary_df["已覆盖比例"].map(lambda x: f"{x:.2f}%")
        ppt_region_summary_df["DHIA认证通过比列"] = ppt_region_summary_df["DHIA认证通过比列"].map(lambda x: f"{x:.2f}%")
        ppt_region_summary_df["Software认证通过比例"] = ppt_region_summary_df["Software认证通过比例"].map(lambda x: f"{x:.2f}%")
        ppt_region_summary_df = ppt_region_summary_df[
            ["大区", "NP客户数量", "已覆盖个数", "已覆盖比例", "DHIA认证通过个数", "DHIA认证通过比列", "Software认证通过个数", "Software认证通过比例"]
        ]

        ppt_subregion_summary_map = {}
        region_order = region_summary["region_l2"].tolist()
        for region_name in region_order:
            region_df = subregion_summary[subregion_summary["region_l2"] == region_name].copy()
            if region_df.empty:
                continue
            ppt_df = region_df.rename(
                columns={
                    "region_l3": "代表处",
                    "np_customer_count": "NP客户数量",
                    "covered_count": "已覆盖个数",
                    "certified_count": "DHIA认证通过个数",
                    "covered_ratio": "已覆盖比例",
                    "certified_ratio": "DHIA认证通过比列",
                }
            )
            ppt_df["已覆盖比例"] = ppt_df["已覆盖比例"].map(lambda x: f"{x:.2f}%")
            ppt_df["DHIA认证通过比列"] = ppt_df["DHIA认证通过比列"].map(lambda x: f"{x:.2f}%")
            ppt_subregion_summary_map[region_name] = ppt_df[
                ["代表处", "NP客户数量", "已覆盖个数", "已覆盖比例", "DHIA认证通过个数", "DHIA认证通过比列"]
            ]

        software_region_summary_df = region_summary[[
            "region_l2",
            "np_customer_count",
            "covered_ratio",
            "software_certified_count",
            "software_certified_ratio",
            "certified_count",
            "certified_ratio",
        ]].copy()
        software_region_summary_df = software_region_summary_df.sort_values(
            by=["software_certified_ratio", "software_certified_count", "region_l2"],
            ascending=[False, False, True],
        ).reset_index(drop=True)
        software_region_summary_df["software_cert_rank"] = software_region_summary_df.index + 1
        software_region_summary_df = software_region_summary_df.rename(
            columns={
                "region_l2": "大区",
                "np_customer_count": "NP客户数量",
                "covered_ratio": "已覆盖比例",
                "software_certified_count": "Software认证通过个数",
                "software_certified_ratio": "Software认证通过比例",
                "software_cert_rank": "Software认证通过比例排名",
                "certified_count": "DHIA认证通过个数",
                "certified_ratio": "DHIA认证通过比例",
            }
        )

        software_subregion_summary_df = subregion_summary[[
            "region_l2",
            "region_l3",
            "np_customer_count",
            "covered_ratio",
            "software_certified_count",
            "software_certified_ratio",
            "certified_count",
            "certified_ratio",
        ]].copy()
        software_subregion_summary_df = software_subregion_summary_df.rename(
            columns={
                "region_l2": "大区",
                "region_l3": "代表处",
                "np_customer_count": "NP客户数量",
                "covered_ratio": "已覆盖比例",
                "software_certified_count": "Software认证通过个数",
                "software_certified_ratio": "Software认证通过比例",
                "certified_count": "DHIA认证通过个数",
                "certified_ratio": "DHIA认证通过比例",
            }
        )

        software_certified_customers_df = detail[detail["software_certified"] == 1].copy()
        if software_certified_customers_df.empty:
            software_certified_customers_df = pd.DataFrame(columns=[
                "客户ID",
                "客户名称",
                "大区",
                "代表处",
                "是否覆盖",
                "是否DHIA认证",
                "Software认证类型",
                "全部证书类型",
                "证书映射依据",
                "是否通过拉美msg覆盖",
            ])
        else:
            software_certified_customers_df = software_certified_customers_df.rename(
                columns={
                    "customer_id": "客户ID",
                    "customer_name": "客户名称",
                    "region_l2": "大区",
                    "region_l3": "代表处",
                    "covered_training": "是否覆盖",
                    "certified": "是否DHIA认证",
                    "software_cert_types": "Software认证类型",
                    "cert_types": "全部证书类型",
                    "cert_mapping_basis": "证书映射依据",
                    "covered_via_latin_msg": "是否通过拉美msg覆盖",
                }
            )
            software_certified_customers_df = software_certified_customers_df[[
                "客户ID",
                "客户名称",
                "大区",
                "代表处",
                "是否覆盖",
                "是否DHIA认证",
                "Software认证类型",
                "全部证书类型",
                "证书映射依据",
                "是否通过拉美msg覆盖",
            ]]

        software_top_uncertified_customers_df = detail[(detail["software_certified"] == 0) & (detail["covered_training"] == 1)].copy()
        if software_top_uncertified_customers_df.empty:
            software_top_uncertified_customers_df = pd.DataFrame(columns=[
                "客户ID",
                "客户名称",
                "大区",
                "代表处",
                "培训活动数",
                "关注课程数",
                "关注课程",
                "是否DHIA认证",
                "全部证书类型",
            ])
        else:
            software_top_uncertified_customers_df = software_top_uncertified_customers_df.sort_values(
                by=["focus_course_count", "training_activity_count", "region_l2", "customer_name"],
                ascending=[False, False, True, True],
            ).head(top_n)
            software_top_uncertified_customers_df = software_top_uncertified_customers_df.rename(
                columns={
                    "customer_id": "客户ID",
                    "customer_name": "客户名称",
                    "region_l2": "大区",
                    "region_l3": "代表处",
                    "training_activity_count": "培训活动数",
                    "focus_course_count": "关注课程数",
                    "focus_courses": "关注课程",
                    "certified": "是否DHIA认证",
                    "cert_types": "全部证书类型",
                }
            )
            software_top_uncertified_customers_df = software_top_uncertified_customers_df[[
                "客户ID",
                "客户名称",
                "大区",
                "代表处",
                "培训活动数",
                "关注课程数",
                "关注课程",
                "是否DHIA认证",
                "全部证书类型",
            ]]

        return {
            "customer_detail_df": detail,
            "region_summary_df": region_summary,
            "region_level3_summary_df": region_level3_summary,
            "subregion_summary_df": subregion_summary,
            "cert_type_count_df": cert_type_count_df,
            "cert_type_rate_df": cert_type_rate_df,
            "coverage_x_cert_df": coverage_x_cert_df,
            "top_uncovered_customers_df": top_uncovered_customers_df,
            "ppt_region_summary_df": ppt_region_summary_df,
            "ppt_subregion_summary_map": ppt_subregion_summary_map,
            "focus_course_activity_df": focus_activity,
            "cert_mapping_used_df": cert_mapping_used_df,
            "software_region_summary_df": software_region_summary_df,
            "software_subregion_summary_df": software_subregion_summary_df,
            "software_certified_customers_df": software_certified_customers_df,
            "software_top_uncertified_customers_df": software_top_uncertified_customers_df,
        }
