import os
import yaml
import pandas as pd
from datetime import datetime
import win32com.client as win32


class CustomerEmailSender:
    def __init__(self, config_path: str):
        self.config_path = config_path
        with open(config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f) or {}
        self.email_cfg = self.config.get("email", {}) or {}
        self.enabled = bool(self.email_cfg.get("enabled", False))
        self.mode = str(self.email_cfg.get("mode", "outlook") or "outlook").strip().lower()
        self.cc_list = str(self.email_cfg.get("cc_list", "") or "").strip()
        self.subject_template = str(self.email_cfg.get("subject_template", "") or "").strip()
        self.group_owners = self.email_cfg.get("group_owners", {}) or {}

    def _sanitize(self, text):
        if text is None:
            return ""
        return " ".join(str(text).replace("\u3000", " ").split()).strip()

    def _fmt_pct(self, v):
        try:
            return f"{float(v):.2f}%"
        except Exception:
            return str(v)

    def _generate_table_html(self, df: pd.DataFrame, columns):
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            return "<p>暂无数据</p>"
        df2 = df.copy().fillna("")
        cols = [c for c in columns if c in df2.columns]
        if not cols:
            return "<p>暂无数据</p>"
        html = "<table><thead><tr>"
        for col in cols:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>"
        for _, row in df2.iterrows():
            html += "<tr>"
            for col in cols:
                html += f"<td>{row.get(col, '')}</td>"
            html += "</tr>"
        html += "</tbody></table>"
        return html

    def _generate_table_html_highlight(self, df: pd.DataFrame, columns, highlight_col: str, highlight_value: str):
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            return "<p>暂无数据</p>"
        df2 = df.copy().fillna("")
        cols = [c for c in columns if c in df2.columns]
        if not cols:
            return "<p>暂无数据</p>"
        highlight_value = self._sanitize(highlight_value)
        html = "<table><thead><tr>"
        for col in cols:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>"
        for _, row in df2.iterrows():
            row_style = ""
            if highlight_col in df2.columns:
                v = self._sanitize(row.get(highlight_col, ""))
                if highlight_value and v == highlight_value:
                    row_style = ' style="background-color: #ffff99; font-weight: bold;"'
            html += f"<tr{row_style}>"
            for col in cols:
                html += f"<td>{row.get(col, '')}</td>"
            html += "</tr>"
        html += "</tbody></table>"
        return html

    def _sanitize_filename(self, name):
        safe = str(name or "")
        for ch in ['\\', '/', '?', '*', '[', ']', ':', '"', '<', '>', '|']:
            safe = safe.replace(ch, " ")
        safe = " ".join(safe.split()).strip()
        return safe or "file"

    def _save_email_preview(self, region_name: str, subject: str, html_body: str):
        out_dir = (self.config.get("paths", {}) or {}).get("output_dir", "outputs/")
        preview_dir = os.path.join(out_dir, "email_previews")
        os.makedirs(preview_dir, exist_ok=True)
        today_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = self._sanitize_filename(f"{region_name}_{today_str}_{subject}")
        path = os.path.join(preview_dir, f"{base}.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html_body)
        return path

    def _attachments_root(self):
        out_dir = (self.config.get("paths", {}) or {}).get("output_dir", "outputs/")
        attach_dir = os.path.join(out_dir, "email_attachments")
        os.makedirs(attach_dir, exist_ok=True)
        return attach_dir

    def _metric_definitions_html(self):
        return """
        <div style="background:#F6F8FA;border:1px solid #E5E7EB;padding:10px 12px;border-radius:6px;margin:0 0 14px 0;">
            <div style="font-weight:700;margin-bottom:6px;">指标定义（口径说明）</div>
            <ol style="margin:0;padding-left:20px;">
                <li><b>NP客户数量</b>：纳入统计的客户数（来自客户清单，过滤 Account Type == NP，剔除 ignore_regions）。</li>
                <li><b>已覆盖个数</b>：该维度下已覆盖客户数量。已覆盖判定=客户参加过至少 1 次关注课程（focus course）对应的培训活动；拉美区额外包含 .msg 邮件正文表格匹配命中的覆盖客户。</li>
                <li><b>已覆盖比例</b>：已覆盖个数 / NP客户数量 × 100%。</li>
                <li><b>DHIA认证通过个数</b>：该维度下已通过 DHIA 认证的客户数（证书清单匹配到该客户）。</li>
                <li><b>DHIA认证通过比例</b>：DHIA认证通过个数 / NP客户数量 × 100%。</li>
                <li><b>Software认证通过个数</b>：证书名称（Certificate Name）包含 “Software”（大小写不敏感）的客户数。</li>
                <li><b>Software认证通过比例</b>：Software认证通过个数 / NP客户数量 × 100%。</li>
            </ol>
        </div>
        """

    def _compute_region_rankings(self, region_df: pd.DataFrame):
        if region_df is None or not isinstance(region_df, pd.DataFrame) or region_df.empty:
            return {}, {}, 0

        df = region_df.copy()
        df["region_l2"] = df["region_l2"].astype(str).apply(self._sanitize)

        df_cov = df.sort_values(
            by=["covered_ratio", "covered_count", "region_l2"],
            ascending=[False, False, True],
        ).reset_index(drop=True)
        df_cov["rank_covered_ratio"] = df_cov.index + 1
        cov_map = dict(zip(df_cov["region_l2"].tolist(), df_cov["rank_covered_ratio"].tolist()))

        df_cert = df.sort_values(
            by=["certified_ratio", "certified_count", "region_l2"],
            ascending=[False, False, True],
        ).reset_index(drop=True)
        df_cert["rank_certified_ratio"] = df_cert.index + 1
        cert_map = dict(zip(df_cert["region_l2"].tolist(), df_cert["rank_certified_ratio"].tolist()))

        return cov_map, cert_map, int(df["region_l2"].nunique())

    def _build_region_rank_tables(self, region_df: pd.DataFrame, highlight_region: str, top_n: int = None):
        if region_df is None or not isinstance(region_df, pd.DataFrame) or region_df.empty:
            return ""

        df = region_df.copy()
        df["大区"] = df["region_l2"].astype(str).apply(self._sanitize)
        df["NP客户数量"] = df.get("np_customer_count", 0).fillna(0).astype(int)
        df["已覆盖比例"] = df.get("covered_ratio", 0).apply(self._fmt_pct)
        df["DHIA认证通过比例"] = df.get("certified_ratio", 0).apply(self._fmt_pct)
        df["Software认证通过比例"] = df.get("software_certified_ratio", 0).apply(self._fmt_pct)

        cov = df.sort_values(
            by=["covered_ratio", "covered_count", "大区"],
            ascending=[False, False, True],
        ).reset_index(drop=True)
        cov["排名"] = cov.index + 1
        cov_disp = cov[["排名", "大区", "NP客户数量", "已覆盖比例", "DHIA认证通过比例", "Software认证通过比例"]].copy()

        if top_n:
            top_n = int(top_n)
            if top_n > 0:
                cov_disp = cov_disp.head(top_n).copy()

        cov_html = self._generate_table_html_highlight(
            cov_disp,
            list(cov_disp.columns),
            highlight_col="大区",
            highlight_value=highlight_region,
        )
        return cov_html

    def _build_uncovered_by_rep(self, detail_df: pd.DataFrame, region_name: str):
        if detail_df is None or not isinstance(detail_df, pd.DataFrame) or detail_df.empty:
            return "", []

        region_name = self._sanitize(region_name)
        d = detail_df.copy()
        if "region_l2" not in d.columns or "covered_training" not in d.columns:
            return "", []

        d["region_l2"] = d["region_l2"].astype(str).apply(self._sanitize)
        d["region_l3"] = d.get("region_l3", "").astype(str).apply(self._sanitize) if "region_l3" in d.columns else ""
        d = d[(d["region_l2"] == region_name) & (d["covered_training"] == 0)].copy()
        if d.empty:
            return "<p><b>未覆盖客户明细：</b>暂无未覆盖客户。</p>", []

        attachments = []
        html_parts = ["<p><b>未覆盖客户明细（按代表处依次列出）：</b></p>"]
        attach_dir = os.path.join(self._attachments_root(), self._sanitize_filename(region_name))
        os.makedirs(attach_dir, exist_ok=True)
        for fn in os.listdir(attach_dir):
            if str(fn).lower().endswith(".xlsx"):
                try:
                    os.remove(os.path.join(attach_dir, fn))
                except Exception:
                    pass

        show_cols_src = []
        for c in [
            "customer_id",
            "customer_name",
            "country_cn",
            "country",
            "account_type",
            "region_l3",
            "region_l3_cn",
            "training_activity_count",
            "focus_course_count",
            "focus_courses",
        ]:
            if c in d.columns:
                show_cols_src.append(c)

        d = d.sort_values(by=["region_l3", "customer_name"] if "customer_name" in d.columns else ["region_l3"]).copy()
        for rep, rep_df in d.groupby("region_l3", dropna=False):
            rep_name = self._sanitize(rep) or "未知"
            rep_cn = ""
            if "region_l3_cn" in rep_df.columns:
                s = rep_df["region_l3_cn"].dropna().astype(str).apply(self._sanitize)
                s = s[s.str.len() > 0]
                if not s.empty:
                    rep_cn = s.iloc[0]
            rep_display = ""
            if rep_cn and rep_cn != rep_name:
                rep_display = f"{rep_cn} ({rep_name})"
            else:
                rep_display = rep_cn or rep_name
            rep_show = rep_df[show_cols_src].copy() if show_cols_src else rep_df.copy()
            if "country_cn" in rep_show.columns and "country" in rep_show.columns:
                rep_show["country_final"] = rep_show["country_cn"].astype(str).apply(self._sanitize)
                rep_show.loc[rep_show["country_final"].str.len() == 0, "country_final"] = rep_show["country"].astype(str).apply(self._sanitize)
            elif "country_cn" in rep_show.columns:
                rep_show["country_final"] = rep_show["country_cn"].astype(str).apply(self._sanitize)
            elif "country" in rep_show.columns:
                rep_show["country_final"] = rep_show["country"].astype(str).apply(self._sanitize)
            else:
                rep_show["country_final"] = ""

            rep_show = rep_show.rename(
                columns={
                    "customer_id": "客户ID",
                    "customer_name": "客户名称",
                    "country_final": "国家",
                    "account_type": "客户类型",
                    "training_activity_count": "关注课程活动次数",
                    "focus_course_count": "关注课程数",
                    "focus_courses": "关注课程清单",
                }
            )
            rep_show = rep_show.drop(columns=[c for c in ["country_cn", "country", "region_l3", "region_l3_cn"] if c in rep_show.columns], errors="ignore")

            rep_file_base = rep_display
            attach_path = os.path.join(
                attach_dir,
                f"{self._sanitize_filename(rep_file_base)}_未覆盖客户清单.xlsx",
            )
            rep_show.to_excel(attach_path, index=False)
            attachments.append(attach_path)

            html_parts.append(f"<div style=\"margin:10px 0 6px 0;\"><b>代表处：{rep_display}</b></div>")
            html_parts.append(self._generate_table_html(rep_show, list(rep_show.columns)))

        return "\n".join(html_parts), attachments

    def build_region_email(
        self,
        region_name: str,
        region_row: dict,
        subregion_df: pd.DataFrame,
        *,
        rank_covered_ratio: int = None,
        rank_certified_ratio: int = None,
        total_regions: int = None,
        rank_table_covered_html: str = "",
        uncovered_html: str = "",
    ):
        region_name = self._sanitize(region_name)
        today_str = datetime.now().strftime("%Y%m%d")
        task_name = self._sanitize(self.config.get("task_name", "客户培训覆盖分析"))

        subj_tpl = self.subject_template or "{group_name} 客户培训覆盖进展情况 {date}"
        subject = subj_tpl.format(group_name=region_name, date=today_str)

        metric_cards = f"""
        <p><b>大区整体情况：</b></p>
        <ul>
            <li>NP客户数量：{int(region_row.get('np_customer_count', 0) or 0)}</li>
            <li>已覆盖：{int(region_row.get('covered_count', 0) or 0)}（{self._fmt_pct(region_row.get('covered_ratio', 0))}）</li>
            <li>DHIA认证通过：{int(region_row.get('certified_count', 0) or 0)}（{self._fmt_pct(region_row.get('certified_ratio', 0))}）</li>
            <li>Software认证通过：{int(region_row.get('software_certified_count', 0) or 0)}（{self._fmt_pct(region_row.get('software_certified_ratio', 0))}）</li>
        </ul>
        """
        rank_lines = []
        if rank_covered_ratio is not None and total_regions:
            rank_lines.append(f"<li>已覆盖比例（所有大区）排名：{int(rank_covered_ratio)}/{int(total_regions)}</li>")
        if rank_certified_ratio is not None and total_regions:
            rank_lines.append(f"<li>DHIA认证通过比例（所有大区）排名：{int(rank_certified_ratio)}/{int(total_regions)}</li>")
        if rank_lines:
            metric_cards += f"<p><b>本大区排名：</b></p><ul>{''.join(rank_lines)}</ul>"

        rank_tables_html = ""
        if rank_table_covered_html:
            rank_tables_html += "<p><b>大区排名明细：</b></p>"
            rank_tables_html += "<p><b>1) 已覆盖比例排名</b></p>"
            rank_tables_html += rank_table_covered_html

        rep_table = pd.DataFrame()
        if subregion_df is not None and isinstance(subregion_df, pd.DataFrame) and not subregion_df.empty:
            rep_table = subregion_df.copy()
            rep_table["已覆盖比例"] = rep_table["covered_ratio"].apply(self._fmt_pct) if "covered_ratio" in rep_table.columns else ""
            rep_table["DHIA认证通过比例"] = rep_table["certified_ratio"].apply(self._fmt_pct) if "certified_ratio" in rep_table.columns else ""
            rep_table["Software认证通过比例"] = rep_table["software_certified_ratio"].apply(self._fmt_pct) if "software_certified_ratio" in rep_table.columns else ""
            rep_table = rep_table.rename(
                columns={
                    "region_l3": "代表处（四级部门）",
                    "np_customer_count": "NP客户数量",
                    "covered_count": "已覆盖个数",
                    "certified_count": "DHIA认证通过个数",
                    "software_certified_count": "Software认证通过个数",
                    "rank_in_region": "大区内排名",
                }
            )
            keep = [
                "大区内排名",
                "代表处（四级部门）",
                "NP客户数量",
                "已覆盖个数",
                "已覆盖比例",
                "DHIA认证通过比例",
                "Software认证通过比例",
            ]
            rep_table = rep_table[[c for c in keep if c in rep_table.columns]].copy()

        rep_html = f"""
        <p><b>代表处（四级部门）明细：</b></p>
        {self._generate_table_html(rep_table, list(rep_table.columns) if not rep_table.empty else ["代表处（四级部门）"])}
        """

        html = f"""
        <html>
        <head>
            <style>
                table {{ border-collapse: collapse; width: 100%; font-family: '微软雅黑', sans-serif; font-size: 14px; margin-bottom: 14px; }}
                th {{ background-color: #f2f2f2; border: 1px solid #dddddd; text-align: left; padding: 8px; }}
                td {{ border: 1px solid #dddddd; text-align: left; padding: 8px; }}
                ul {{ margin-top: 6px; }}
            </style>
        </head>
        <body>
            <p>你好，</p>
            <p>以下为 <b>{region_name}</b> 的 <b>{task_name}</b> 最新进展同步：</p>
            {self._metric_definitions_html()}
            {metric_cards}
            {rank_tables_html}
            {rep_html}
            {uncovered_html or ""}
            <p>祝好！</p>
        </body>
        </html>
        """
        return subject, html

    def send_email(
        self,
        to_emails: str,
        subject: str,
        html_body: str,
        cc_emails: str = "",
        subject_prefix: str = None,
        attachments=None,
    ):
        if self.mode != "outlook":
            raise ValueError(f"仅支持 outlook 模式，当前 mode={self.mode}")
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.Subject = f"{subject_prefix}{subject}" if subject_prefix else subject
        mail.To = str(to_emails or "").replace(",", ";")
        mail.CC = str(cc_emails or "").replace(",", ";")
        mail.HTMLBody = html_body
        for p in (attachments or []):
            if not p:
                continue
            ap = os.path.abspath(str(p))
            if os.path.exists(ap):
                mail.Attachments.Add(ap)
        mail.Send()
        return True

    def send_customer_coverage_notifications(
        self,
        dfs: dict,
        region_filter=None,
        test_email_only: bool = False,
        test_recipient: str = "",
        subject_prefix: str = "【测试】",
        disable_cc: bool = False,
        preview_only: bool = False,
        force_send: bool = False,
    ):
        if (not self.enabled) and (not test_email_only) and (not force_send):
            print("ℹ️ 邮件功能未开启（email.enabled=false），跳过发送。")
            return
        if (not self.enabled) and (not test_email_only) and force_send:
            print("⚠️ 邮件功能未开启（email.enabled=false），但已指定 --force-email，仍将尝试发送。")

        region_df = dfs.get("region_summary_df")
        subregion_df = dfs.get("subregion_summary_df")
        detail_df = dfs.get("customer_detail_df")
        if region_df is None or not isinstance(region_df, pd.DataFrame) or region_df.empty:
            print("⚠️ 缺少 region_summary_df，无法发送客户覆盖邮件。")
            return

        cov_rank_map, cert_rank_map, total_regions = self._compute_region_rankings(region_df)

        targets = region_df.copy()
        if region_filter:
            region_set = set([self._sanitize(x) for x in region_filter if self._sanitize(x)])
            targets = targets[targets["region_l2"].astype(str).apply(self._sanitize).isin(region_set)].copy()

        for _, row in targets.iterrows():
            region_name = self._sanitize(row.get("region_l2", ""))
            if not region_name:
                continue

            owner = self.group_owners.get(region_name) or {}
            to_email = ""
            if test_email_only:
                to_email = self._sanitize(test_recipient)
            else:
                to_email = self._sanitize(owner.get("email", ""))

            if not to_email:
                print(f"⚠️ 未找到收件人邮箱，跳过发送：{region_name}")
                continue

            region_row = row.to_dict()
            sub = pd.DataFrame()
            if subregion_df is not None and isinstance(subregion_df, pd.DataFrame) and not subregion_df.empty:
                sub = subregion_df[subregion_df["region_l2"].astype(str).apply(self._sanitize) == region_name].copy()

            uncovered_html, attachments = self._build_uncovered_by_rep(detail_df, region_name)
            rank_cov_html = self._build_region_rank_tables(region_df, highlight_region=region_name)
            subject, html = self.build_region_email(
                region_name,
                region_row,
                sub,
                rank_covered_ratio=cov_rank_map.get(region_name),
                rank_certified_ratio=cert_rank_map.get(region_name),
                total_regions=total_regions,
                rank_table_covered_html=rank_cov_html,
                uncovered_html=uncovered_html,
            )
            if preview_only:
                preview_path = self._save_email_preview(region_name, subject, html)
                print(f"📝 已生成邮件预览：{region_name} -> {preview_path}")
                if attachments:
                    print(f"📎 已生成附件({len(attachments)}个)：{region_name} -> {os.path.dirname(attachments[0])}")
                continue
            cc_emails = "" if (disable_cc or test_email_only) else self.cc_list
            prefix = subject_prefix if test_email_only else None
            self.send_email(
                to_email,
                subject,
                html,
                cc_emails=cc_emails,
                subject_prefix=prefix,
                attachments=attachments,
            )
            print(f"✅ 已发送客户覆盖邮件：{region_name} -> {to_email}")
