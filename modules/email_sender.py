import os
import shutil
import yaml
import pandas as pd
from datetime import datetime
import smtplib
import win32com.client as win32
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class EmailSender:
    def __init__(self, config_path="config/config.yaml"):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        self.email_config = self.config.get('email', {})
        self.enabled = self.email_config.get('enabled', False)
        self.mode = self.email_config.get('mode', 'outlook') # 'outlook' or 'smtp'
        self.cc_list = self.email_config.get('cc_list', '')
        self.group_owners = self.email_config.get('group_owners', {})
        self._group_name_en_map = {
            "非洲区": "Africa",
            "中东北非区": "MENA",
            "欧亚区": "Eurasia",
            "大洋洲区": "Oceania",
            "英国区": "UK",
            "南洋区": "Nanyang",
            "东北欧区": "Northeast Europe",
            "拉美区": "Latin America",
            "东南亚区": "Southeast Asia",
            "巴西区": "Brazil",
            "墨西哥区": "Mexico",
            "西欧区": "Western Europe",
            "国际合作部": "International Cooperation Dept",
            "集成商业务部": "Integrator Business Dept",
            "产品市场部": "Product Marketing Dept",
            "总部": "HQ",
        }
        self._brazil_level4_en_map = {
            "解决方案部": "Solutions Dept",
            "项目业务部": "Project Business Dept",
        }

    def _to_group_name_en(self, name):
        key = str(name).replace(" ", "").replace("\n", "").strip()
        return self._group_name_en_map.get(key, key)

    def _to_task_name_en(self, task_name):
        name = str(task_name or "").strip()
        if "DHIA Software" in name:
            return "DHIA Software Learning"
        if "必知必会" in name:
            return "Essential Skills Learning"
        return name

    def _bilingual_attachment_filename(self, group_name, cn_suffix, en_suffix, level4_name=None):
        group_cn = str(group_name or "").strip()
        group_en = self._to_group_name_en(group_cn)
        if level4_name is not None and str(level4_name).strip():
            level4_cn = str(level4_name).strip()
            level4_en = self._translate_brazil_text_en(level4_cn) if group_cn == "巴西区" else level4_cn
            left = f"{group_cn}-{level4_cn}{cn_suffix}"
            right = f"{group_en}-{level4_en}{en_suffix}"
        else:
            left = f"{group_cn}{cn_suffix}"
            right = f"{group_en}{en_suffix}"
        return self._sanitize_filename(f"{left} - {right}.xlsx")

    def _deadline_reminder_html(self, task_name, include_cn=True, include_en=True):
        name = str(task_name or "").strip()
        cn = ""
        en = ""
        if "DHIA Software" in name:
            cn = "软件DHIA学习任务截至时间为6月20日，请尽快推荐完成。"
            en = "The deadline for the DHIA Software Learning task is June 20. Please expedite completion."
        elif "必知必会" in name or "极客" in name:
            cn = "极客必知必会学习任务截至时间为 6月20日，请尽快推进完成。"
            en = "The deadline for the Essential Skills Learning task is June 20. Please expedite completion."
        else:
            return ""

        parts = []
        if include_cn and cn:
            parts.append(f"<div><b>{cn}</b></div>")
        if include_en and en:
            parts.append(f"<div><b>{en}</b></div>")
        if not parts:
            return ""
        return f"""
        <div style="background:#FFF3CD;border:1px solid #FFEEBA;padding:10px 12px;border-radius:6px;margin:0 0 14px 0;color:#856404;">
            {''.join(parts)}
        </div>
        """

    def _translate_brazil_text_en(self, text):
        if text is None:
            return ""
        s = str(text)
        replacements = [
            ("巴西区", "Brazil"),
            ("解决方案部", self._brazil_level4_en_map.get("解决方案部", "Solutions Dept")),
            ("项目业务部", self._brazil_level4_en_map.get("项目业务部", "Project Business Dept")),
            ("海外", "Overseas"),
            ("代表处", "Rep Office"),
            ("->", " > "),
        ]
        for a, b in replacements:
            s = s.replace(a, b)
        return s

    def _build_rank_table_html_en(self, rank_df_display, highlight_group_name_cn):
        df = rank_df_display.copy()
        if "排名" not in df.columns and "本轮排名" in df.columns:
            df["排名"] = df["本轮排名"]
        cols_cn = ["排名", "小组", "学习完成率", "上一轮排名", "上一轮学习完成率"]
        existing = [c for c in cols_cn if c in df.columns]
        df = df[existing]
        rename_map = {
            "排名": "Rank",
            "小组": "Group",
            "学习完成率": "Completion Rate",
            "上一轮排名": "Prev Rank",
            "上一轮学习完成率": "Prev Completion",
        }
        df = df.rename(columns=rename_map)
        if "Group" in df.columns:
            df["Group"] = df["Group"].apply(self._to_group_name_en)
        cols_en = [rename_map[c] for c in existing if c in rename_map]
        highlight_en = self._to_group_name_en(highlight_group_name_cn)
        return self._generate_table_html(df, cols_en, highlight_group=highlight_en)

    def _sanitize_filename(self, name):
        safe = str(name)
        for ch in ['\\', '/', '?', '*', '[', ']', ':', '"', '<', '>', '|']:
            safe = safe.replace(ch, ' ')
        safe = safe.strip()
        if not safe:
            safe = "file"
        return safe

    def _sanitize_sheet_name(self, name):
        safe = str(name)
        for ch in ['\\', '/', '?', '*', '[', ']', ':']:
            safe = safe.replace(ch, ' ')
        safe = safe.strip()
        if not safe:
            safe = "Sheet"
        if len(safe) > 31:
            safe = safe[:31]
        return safe

    def _attachment_tag(self):
        task_name = str(self.config.get('task_name', '') or '')
        if 'DHIA Software' in task_name:
            return 'DHIA Software'
        return ""

    def _extract_person_columns(self, df):
        col_map = {}
        for target in ['账号', '姓名', '部门', '岗位']:
            for col in df.columns:
                if target in str(col):
                    col_map[target] = col
                    break
        cols = [c for c in col_map.values() if c]
        return col_map, cols
    
    def _normalize_col_name(self, col):
        return "".join(str(col).replace("\u3000", " ").split()).strip()
    
    def _detect_status_col(self, df):
        candidates = []
        for col in df.columns:
            norm = self._normalize_col_name(col)
            if not norm:
                continue
            if norm in {"完成状态", "学习状态"}:
                return col
            if "完成" in norm and "状态" in norm:
                candidates.append(col)
            elif "学习" in norm and "状态" in norm:
                candidates.append(col)
        return candidates[0] if candidates else df.columns[-1]

    def _build_person_table_df(self, df, col_map, include_position=True):
        out = pd.DataFrame()
        targets = ['账号', '姓名', '部门', '岗位'] if include_position else ['账号', '姓名', '部门']
        for target in targets:
            src = col_map.get(target)
            out[target] = df[src] if src in df.columns else ""
        return out

    def _build_level4_summary(self, df, status_col):
        grouped = df.groupby('四级部门').agg(
            总人数=(status_col, 'count'),
            已完成人数=(status_col, lambda x: (x == '已完成').sum())
        ).reset_index().rename(columns={'四级部门': '小组'})
        if not grouped.empty:
            grouped['学习完成率'] = (grouped['已完成人数'] / grouped['总人数'] * 100).round(2)
        return grouped

    def send_group_notification(self, group_name, stats, attachment_path, rank_table_html, completed_table_html, uncompleted_table_html, rep_detail_html=None, extra_attachment_paths=None, override_recipient=None, disable_cc=False, subject_prefix=None, rank_table_html_en=None, rep_detail_html_en=None):
        """通过本地 Outlook 发送小组通知邮件"""
        if not self.enabled:
            print(f"ℹ️ 邮件功能未开启，跳过对 {group_name} 的通知。")
            return False

        owner_info = self.group_owners.get(group_name)
        if not owner_info:
            print(f"⚠️ 未找到小组 '{group_name}' 的负责人配置，跳过发送。")
            return False

        # 如果提供了 override_recipient (测试模式)，则发送到指定地址
        recipient_email = override_recipient if override_recipient else owner_info.get('email')
        owner_name = owner_info.get('name', '负责人')

        try:
            # 1. 启动并连接 Outlook
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0) # 0 代表 olMailItem

            # 2. 设置邮件基本信息
            today_str = datetime.now().strftime("%Y%m%d")
            task_name = self.config.get('task_name', '2026 极客行动必知必会学习')
            if group_name == "巴西区":
                group_name_en = self._to_group_name_en(group_name)
                task_name_en = self._to_task_name_en(task_name)
                subject = f"{task_name_en} Progress Update | {group_name_en} | {today_str} | Completion {stats['学习完成率']}%"
            else:
                subject = self.email_config.get('subject_template', "").format(
                    group_name=group_name,
                    completion_rate=stats['学习完成率'],
                    date=today_str
                )
            if subject_prefix:
                subject = f"{subject_prefix}{subject}"
            mail.Subject = subject
            # Outlook 的收件人如果有多位，通常需要用分号 ; 分隔
            mail.To = recipient_email.replace(',', ';')
            
            # 抄送 (CC) 逻辑
            if (not disable_cc) and self.cc_list:
                mail.CC = self.cc_list.replace(',', ';')
            else:
                mail.CC = ""

            # 3. 准备 HTML 正文
            rank_change_desc = ""
            curr_rank = stats.get('本轮排名')
            prev_rank = stats.get('上一轮排名')
            
            def safe_int(val):
                if val is None:
                    return None
                s = str(val).strip()
                if not s or s == '-' or s.lower() == 'nan':
                    return None
                try:
                    return int(float(s))
                except Exception:
                    return None
            
            curr_rank_num = safe_int(curr_rank)
            prev_rank_num = safe_int(prev_rank)
            if prev_rank_num is not None and curr_rank_num is not None:
                try:
                    diff = prev_rank_num - curr_rank_num
                    if diff > 0:
                        rank_change_desc = f"上升了 {diff} 名"
                    elif diff < 0:
                        rank_change_desc = f"下降了 {abs(diff)} 名"
                    else:
                        rank_change_desc = "保持不变"
                except:
                    rank_change_desc = "本轮为首次对比分析"
            else:
                rank_change_desc = "本轮为首次对比分析"

            rank_change_desc_en = ""
            if prev_rank_num is not None and curr_rank_num is not None:
                try:
                    diff = prev_rank_num - curr_rank_num
                    if diff > 0:
                        rank_change_desc_en = f"improved by {diff}"
                    elif diff < 0:
                        rank_change_desc_en = f"dropped by {abs(diff)}"
                    else:
                        rank_change_desc_en = "remained unchanged"
                except Exception:
                    rank_change_desc_en = "this is the first comparison"
            else:
                rank_change_desc_en = "this is the first comparison"

            rep_detail_html = rep_detail_html or ""

            incomplete_cnt = int(stats.get('总人数', 0) or 0) - int(stats.get('已完成人数', 0) or 0)
            has_extra_attachments = any([p and os.path.exists(p) for p in (extra_attachment_paths or [])])
            has_uncompleted_attachment = group_name != "总部" and attachment_path and os.path.exists(attachment_path)
            has_attachments = group_name != "总部" and (has_uncompleted_attachment or has_extra_attachments)
            attachment_tip = ""
            if has_uncompleted_attachment:
                attachment_tip = "<p>如需下载明细，请查看邮件附件中的“未完成学员名单”表格。请协助督促未完成学员尽快完成学习任务。</p>"
            elif incomplete_cnt <= 0:
                if has_attachments:
                    attachment_tip = "<p>本小组已全部完成，无未完成学员。明细请查看附件中的“已完成学员名单”。</p>"
                else:
                    attachment_tip = "<p>本小组已全部完成，无未完成学员。</p>"
            else:
                if has_attachments:
                    attachment_tip = "<p>如需下载明细，请查看邮件附件。请协助督促未完成学员尽快完成学习任务。</p>"
                else:
                    attachment_tip = "<p>请协助督促未完成学员尽快完成学习任务。</p>"

            attachment_tip_en = ""
            if has_uncompleted_attachment:
                attachment_tip_en = "<p>Please download the details from the attachment \"Uncompleted Learners List\" and kindly follow up with the incomplete learners to finish the learning tasks as soon as possible.</p>"
            elif incomplete_cnt <= 0:
                if has_attachments:
                    attachment_tip_en = "<p>All learners in this group have completed the learning tasks. Please refer to the attachment \"Completed Learners List\" for details.</p>"
                else:
                    attachment_tip_en = "<p>All learners in this group have completed the learning tasks.</p>"
            else:
                if has_attachments:
                    attachment_tip_en = "<p>Please refer to the email attachments for details, and kindly follow up with the incomplete learners to finish the learning tasks as soon as possible.</p>"
                else:
                    attachment_tip_en = "<p>Please kindly follow up with the incomplete learners to finish the learning tasks as soon as possible.</p>"

            extra_english_section = ""
            if group_name == "巴西区":
                group_name_en = self._to_group_name_en(group_name)
                task_name_en = self._to_task_name_en(task_name)
                rank_table_html_for_en = rank_table_html_en if rank_table_html_en else rank_table_html
                rep_detail_html_for_en = rep_detail_html_en or ""
                reminder_en_only = self._deadline_reminder_html(task_name, include_cn=False, include_en=True)
                extra_english_section = f"""
                {reminder_en_only}
                <p>Dear {owner_name},</p>
                <p>Below is the latest learning progress update for <b>{group_name_en}</b> in <b>{task_name_en}</b>:</p>

                <p><b>1. Completion status:</b><br>
                Total learners: {stats['总人数']}; Completed: {stats['已完成人数']}; Incomplete: {stats['总人数'] - stats['已完成人数']}; <b>Completion rate: {stats['学习完成率']}%</b>.</p>

                <p><b>2. Ranking:</b><br>
                The current rank is <b>No.{curr_rank}</b> among all groups. Details are as follows:</p>
                {rank_table_html_for_en}

                <p><b>3. Rank change:</b><br>
                Compared with the previous cycle, the rank has <b>{rank_change_desc_en}</b>.</p>

                {rep_detail_html_for_en}

                {attachment_tip_en}
                <p>Best regards,</p>
                """

            if group_name == "巴西区":
                body_main = extra_english_section
            else:
                reminder_bilingual = self._deadline_reminder_html(task_name, include_cn=True, include_en=True)
                body_main = f"""
                {reminder_bilingual}
                <p>亲爱的 {owner_name}，你好：</p>
                <p>对于 <b>{group_name}</b> 的 {task_name}进展情况如下：</p>
                
                <p><b>1. 本小组完成情况：</b><br>
                截止目前，本小组总人数 {stats['总人数']} 人，已完成 {stats['已完成人数']} 人，未完成 {stats['总人数'] - stats['已完成人数']} 人，<b>整体完成率为 {stats['学习完成率']}%</b>。</p>
                
                <p><b>2. 排名情况：</b><br>
                目前本小组在所有小组中排名第 <b>{curr_rank}</b> 名。以下是详细排名情况：</p>
                {rank_table_html}
                
                <p><b>3. 名次变化情况：</b><br>
                相比上一轮分析，本小组名次 <b>{rank_change_desc}</b>。</p>
                {rep_detail_html}
                
                {attachment_tip}
                
                <p>祝好！</p>
                """

            html_content = f"""
            <html>
            <head>
                <style>
                    table {{ border-collapse: collapse; width: 100%; font-family: '微软雅黑', sans-serif; font-size: 14px; margin-bottom: 20px; }}
                    th {{ background-color: #f2f2f2; border: 1px solid #dddddd; text-align: left; padding: 8px; }}
                    td {{ border: 1px solid #dddddd; text-align: left; padding: 8px; }}
                </style>
            </head>
            <body>
                {body_main}
            </body>
            </html>
            """
            mail.HTMLBody = html_content

            # 4. 添加附件 (针对“总部”小组不添加附件)
            if group_name != "总部" and attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(os.path.abspath(attachment_path))

            if group_name != "总部" and extra_attachment_paths:
                for p in extra_attachment_paths:
                    if p and os.path.exists(p):
                        mail.Attachments.Add(os.path.abspath(p))

            # 5. 发送邮件
            mail.Send()
            print(f"✅ 已通过 Outlook 向 {group_name} ({recipient_email}) 发送进度通知。")
            return True

        except Exception as e:
            print(f"❌ 通过 Outlook 发送邮件失败: {e}")
            return False

    def _generate_table_html(self, df, columns, highlight_group=None):
        """生成 HTML 表格代码，支持对指定小组行进行高亮"""
        if df.empty:
            return "<p>暂无数据</p>"
        
        html = "<table><thead><tr>"
        for col in columns:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>"
        
        for _, row in df.iterrows():
            # 检查是否需要高亮（针对排名表）
            row_style = ""
            row_group = str(row.get('小组', row.get('Group', ''))).strip()
            if highlight_group and row_group == str(highlight_group).strip():
                row_style = ' style="background-color: #ffff99; font-weight: bold;"' # 浅黄色背景+加粗
            
            html += f"<tr{row_style}>"
            for col in columns:
                val = row.get(col, "")
                html += f"<td>{val}</td>"
            html += "</tr>"
        
        html += "</tbody></table>"
        return html

    def send_all_notifications(self, final_df, uncompleted_dir, all_raw_dfs, test_mode=False, rep_mode=False, group_filter=None, exclude_filter=None, test_email_only=False, test_recipient="morgan.hu@dahuatech.com"):
        """遍历所有小组并发送通知"""
        if not self.enabled:
            print("\nℹ️ 邮件发送功能未开启，请在 config.yaml 中配置 enabled: true。")
            return

        print("\n[邮件阶段] 正在通过本地 Outlook 发送小组进度通知...")
        
        # 合并原始明细数据
        combined_raw = pd.concat(all_raw_dfs, ignore_index=True)
        status_col = self._detect_status_col(combined_raw)

        # 准备排名基础数据 (不在此处生成 HTML，因为每封邮件的高亮行不同)
        rank_cols = ['排名', '小组', '学习完成率', '上一轮排名', '上一轮学习完成率']
        rank_df_display = final_df[final_df['小组'] != '总计'].copy()
        rank_df_display['排名'] = rank_df_display['本轮排名']

        # 2. 遍历发送
        groups_df = final_df[final_df['小组'] != '总计']
        
        col_map, _ = self._extract_person_columns(combined_raw)

        base_output_dir = self.config.get('paths', {}).get('output_dir', 'outputs/')
        rep_output_dir = os.path.join(base_output_dir, "uncompleted_lists_level4")
        os.makedirs(rep_output_dir, exist_ok=True)

        prev_path = self.config.get('paths', {}).get('previous_analysis')
        base_output_dir = self.config.get('paths', {}).get('output_dir', 'outputs/')
        email_attach_dir = os.path.join(base_output_dir, "email_attachments")
        os.makedirs(email_attach_dir, exist_ok=True)
        completed_output_dir = os.path.join(base_output_dir, "completed_lists")
        os.makedirs(completed_output_dir, exist_ok=True)
        attach_tag = self._attachment_tag()

        for _, row in groups_df.iterrows():
            group_name = row['小组']

            if group_filter and group_name not in group_filter:
                continue
            if exclude_filter and group_name in exclude_filter:
                continue
            
            # 如果是测试模式，只发送“中东北非区”
            if (not test_email_only) and test_mode and group_name != '中东北非区':
                continue

            # 1. 为当前邮件生成带高亮的排名表 HTML
            rank_table_html = self._generate_table_html(rank_df_display, rank_cols, highlight_group=group_name)
            rank_table_html_en = None
            if group_name == "巴西区":
                rank_table_html_en = self._build_rank_table_html_en(rank_df_display, highlight_group_name_cn=group_name)

            group_raw = combined_raw[combined_raw['分析部门'] == group_name]
            completed_html = ""
            uncompleted_html = ""
            completed_attachment_path = None
            if group_name != "总部":
                try:
                    completed_df = group_raw[group_raw[status_col] == '已完成']
                    completed_person_df = self._build_person_table_df(completed_df, col_map, include_position=False)
                    prefix = f"{attach_tag}-" if attach_tag else ""
                    file_name = f"{prefix}{self._bilingual_attachment_filename(group_name, '已完成学员清单', ' Completed Learners List')}"
                    completed_attachment_path = os.path.join(completed_output_dir, file_name)
                    completed_person_df.to_excel(completed_attachment_path, index=False)
                except Exception:
                    completed_attachment_path = None

            attachment_path = ""
            if group_name != "总部":
                try:
                    uncompleted_df = group_raw[group_raw[status_col] != "已完成"].copy()
                    if not uncompleted_df.empty:
                        prefix = f"{attach_tag}-" if attach_tag else ""
                        tagged_name = f"{prefix}{self._bilingual_attachment_filename(group_name, '未完成学员清单', ' Uncompleted Learners List')}"
                        tagged_path = os.path.join(email_attach_dir, tagged_name)
                        drop_cols = [c for c in uncompleted_df.columns if c == "分析部门" or ("岗位" in str(c)) or ("Position" in str(c))]
                        if drop_cols:
                            uncompleted_df = uncompleted_df.drop(columns=[c for c in drop_cols if c in uncompleted_df.columns])
                        uncompleted_df.to_excel(tagged_path, index=False)
                        attachment_path = tagged_path
                except Exception:
                    attachment_path = ""
            
            rep_detail_html = ""
            rep_detail_html_en = ""
            extra_attachment_paths = []
            if completed_attachment_path and os.path.exists(completed_attachment_path):
                extra_attachment_paths.append(completed_attachment_path)
            if group_name != "总部" and '四级部门' in group_raw.columns:
                group_rep_dir = os.path.join(rep_output_dir, self._sanitize_filename(group_name))
                os.makedirs(group_rep_dir, exist_ok=True)
                prefix = f"{attach_tag}-" if attach_tag else ""
                for level4_name, level4_df in group_raw.groupby('四级部门'):
                    level4_uncompleted = level4_df[level4_df[status_col] != '已完成']
                    level4_person_df = self._build_person_table_df(level4_uncompleted, col_map, include_position=False)
                    if level4_person_df.empty:
                        continue
                    safe_name = f"{prefix}{self._bilingual_attachment_filename(group_name, '未完成学员清单', ' Uncompleted Learners List', level4_name=level4_name)}"
                    file_path = os.path.join(group_rep_dir, safe_name)
                    try:
                        level4_person_df.to_excel(file_path, index=False)
                        extra_attachment_paths.append(file_path)
                    except Exception:
                        pass

            force_level4_email_content = True
            if force_level4_email_content and '四级部门' in combined_raw.columns and '三级部门' in combined_raw.columns:
                group_level3_raw = combined_raw[combined_raw['三级部门'] == group_name]
                if not group_level3_raw.empty:
                    level4_base = self._build_level4_summary(group_level3_raw, status_col)
                    prev_sheet_df = None
                    if prev_path and os.path.exists(prev_path):
                        try:
                            prev_sheet_df = pd.read_excel(prev_path, sheet_name=self._sanitize_sheet_name(group_name))
                        except Exception:
                            prev_sheet_df = None

                    level4_final = None
                    try:
                        level4_final = self._merge_like(level4_base, prev_sheet_df)
                    except Exception:
                        level4_final = level4_base

                    if level4_final is not None and not level4_final.empty:
                        level4_display = level4_final.copy()
                        if '小组' in level4_display.columns:
                            level4_display = level4_display.rename(columns={'小组': '四级部门'})
                        level4_rank_html = self._generate_table_html(
                            level4_display,
                            ['本轮排名', '四级部门', '学习完成率', '上一轮排名', '上一轮学习完成率', '总人数', '已完成人数']
                        )
                    else:
                        level4_rank_html = "<p>暂无数据</p>"

                    rep_detail_parts = []
                    rep_detail_parts.append("<p><b>4. 四级部门排名情况：</b></p>")
                    rep_detail_parts.append(level4_rank_html)
                    rep_detail_parts.append("<p><b>5. 未完成学员名单（按四级部门划分）：</b></p>")

                    has_any_uncompleted = False
                    for level4_name, level4_df in group_level3_raw.groupby('四级部门'):
                        level4_uncompleted = level4_df[level4_df[status_col] != '已完成']
                        level4_person_df = self._build_person_table_df(level4_uncompleted, col_map, include_position=False)
                        if level4_person_df.empty:
                            continue
                        has_any_uncompleted = True
                        rep_detail_parts.append(f"<p><b>四级部门：{level4_name}（未完成 {len(level4_person_df)} 人）</b></p>")
                        rep_detail_parts.append(self._generate_table_html(level4_person_df, ['账号', '姓名', '部门']))

                    if not has_any_uncompleted:
                        rep_detail_parts.append("<p>本小组各四级部门均已完成，无未完成学员。</p>")

                    rep_detail_html = "\n".join(rep_detail_parts)

                    if group_name == "巴西区":
                        level4_rank_html_en = "<p>No data</p>"
                        if level4_final is not None and not level4_final.empty:
                            level4_en = level4_final.copy()
                            if "小组" in level4_en.columns:
                                level4_en = level4_en.rename(columns={"小组": "Level-4 Dept"})
                            elif "四级部门" in level4_en.columns:
                                level4_en = level4_en.rename(columns={"四级部门": "Level-4 Dept"})
                            rename_cols = {
                                "本轮排名": "Rank",
                                "Level-4 Dept": "Level-4 Dept",
                                "学习完成率": "Completion Rate",
                                "上一轮排名": "Prev Rank",
                                "上一轮学习完成率": "Prev Completion",
                                "总人数": "Total",
                                "已完成人数": "Completed",
                            }
                            level4_en = level4_en.rename(columns=rename_cols)
                            if "Level-4 Dept" in level4_en.columns:
                                level4_en["Level-4 Dept"] = level4_en["Level-4 Dept"].apply(self._translate_brazil_text_en)
                            cols = [c for c in ["Rank", "Level-4 Dept", "Completion Rate", "Prev Rank", "Prev Completion", "Total", "Completed"] if c in level4_en.columns]
                            level4_rank_html_en = self._generate_table_html(level4_en, cols)

                        rep_detail_parts_en = []
                        rep_detail_parts_en.append("<p><b>4. Level-4 department ranking:</b></p>")
                        rep_detail_parts_en.append(level4_rank_html_en)
                        rep_detail_parts_en.append("<p><b>5. Uncompleted learners list (by Level-4 department):</b></p>")

                        has_any_uncompleted_en = False
                        for level4_name, level4_df in group_level3_raw.groupby("四级部门"):
                            level4_uncompleted = level4_df[level4_df[status_col] != "已完成"]
                            level4_person_df = self._build_person_table_df(level4_uncompleted, col_map, include_position=False)
                            if level4_person_df.empty:
                                continue
                            has_any_uncompleted_en = True
                            level4_name_en = self._translate_brazil_text_en(level4_name)
                            rep_detail_parts_en.append(
                                f"<p><b>Level-4 Dept: {level4_name_en} (Incomplete {len(level4_person_df)} learners)</b></p>"
                            )
                            level4_person_en = level4_person_df.rename(
                                columns={"账号": "Account", "姓名": "Name", "部门": "Department"}
                            )
                            if "Department" in level4_person_en.columns:
                                level4_person_en["Department"] = level4_person_en["Department"].apply(self._translate_brazil_text_en)
                            rep_detail_parts_en.append(self._generate_table_html(level4_person_en, ["Account", "Name", "Department"]))

                        if not has_any_uncompleted_en:
                            rep_detail_parts_en.append(
                                "<p>All Level-4 departments have completed the learning tasks. No incomplete learners.</p>"
                            )

                        rep_detail_html_en = "\n".join(rep_detail_parts_en)

            rate_str = str(row['学习完成率']).replace('%', '')
            try:
                rate = float(rate_str)
            except:
                rate = 0.0

            stats = {
                '小组': group_name,
                '总人数': int(row['总人数']),
                '已完成人数': int(row['已完成人数']),
                '学习完成率': rate,
                '本轮排名': row['本轮排名'],
                '上一轮排名': row['上一轮排名']
            }
            
            # 测试模式下，覆盖收件人为总部负责人
            override_recipient = None
            disable_cc = False
            subject_prefix = None
            if test_email_only:
                override_recipient = test_recipient
                disable_cc = True
                subject_prefix = "【测试】"
                print(f"🧪 [测试邮件] 正在将 {group_name} 的内容发送给: {override_recipient}")
            elif test_mode:
                # 发送到总部的邮箱进行确认
                override_recipient = self.group_owners.get('总部', {}).get('email')
                disable_cc = True
                print(f"🧪 [测试模式] 正在将 {group_name} 的内容发送给测试账号: {override_recipient}")

            self.send_group_notification(
                group_name, stats, attachment_path, 
                rank_table_html, completed_html, uncompleted_html,
                rep_detail_html=rep_detail_html,
                extra_attachment_paths=extra_attachment_paths,
                override_recipient=override_recipient,
                disable_cc=disable_cc,
                subject_prefix=subject_prefix,
                rank_table_html_en=rank_table_html_en,
                rep_detail_html_en=rep_detail_html_en
            )

    def _merge_like(self, current_grouped, previous_df):
        final_grouped = current_grouped.copy()
        if final_grouped.empty:
            out = final_grouped.copy()
            out['本轮排名'] = []
            out['上一轮排名'] = []
            out['上一轮学习完成率'] = []
            return out

        if '学习完成率' not in final_grouped.columns:
            final_grouped['学习完成率'] = (final_grouped['已完成人数'] / final_grouped['总人数'] * 100).round(2)

        final_grouped = final_grouped.sort_values(by='学习完成率', ascending=False).reset_index(drop=True)
        final_grouped['本轮排名'] = final_grouped.index + 1
        final_grouped['上一轮排名'] = ""
        final_grouped['上一轮学习完成率'] = ""

        if previous_df is not None and isinstance(previous_df, pd.DataFrame) and not previous_df.empty and '小组' in previous_df.columns:
            prev_data = previous_df[previous_df['小组'] != '总计'].copy() if '总计' in previous_df['小组'].astype(str).values else previous_df.copy()
            prev_data['小组_key'] = prev_data['小组'].astype(str).str.replace(' ', '').str.replace('\n', '')
            rank_map = dict(zip(prev_data['小组_key'], prev_data.get('本轮排名', "")))
            rate_map = dict(zip(prev_data['小组_key'], prev_data.get('学习完成率', "")))

            def get_prev_rank(name):
                key = str(name).replace(' ', '').replace('\n', '')
                return rank_map.get(key, "")

            def get_prev_rate(name):
                key = str(name).replace(' ', '').replace('\n', '')
                return rate_map.get(key, "")

            final_grouped['上一轮排名'] = final_grouped['小组'].apply(get_prev_rank)
            final_grouped['上一轮学习完成率'] = final_grouped['小组'].apply(get_prev_rate)

        out = final_grouped.copy()
        out['学习完成率'] = out['学习完成率'].apply(lambda x: f"{x}%" if x != '-' else x)
        cols = ['小组', '总人数', '已完成人数', '学习完成率', '本轮排名', '上一轮排名', '上一轮学习完成率']
        out = out[cols]
        return out
