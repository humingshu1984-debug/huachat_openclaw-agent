import os
import yaml
import sys
import pandas as pd
from modules.data_analyzer import DataAnalyzer
from reports.ppt_generator import PPTGenerator

def _configure_stdio():
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(errors="replace")
    except Exception:
        pass

def launch_gui():
    _configure_stdio()
    import subprocess
    import threading
    import queue
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox
    from tkinter import filedialog
    from tkinter.scrolledtext import ScrolledText

    repo_root = os.path.dirname(os.path.abspath(__file__))

    def load_groups():
        try:
            analyzer = DataAnalyzer("config/config.yaml")
            groups = list(getattr(analyzer, "template_groups", []) or [])
            groups = [g for g in groups if g and str(g).strip() and str(g).strip() != "总计"]
            return groups
        except Exception:
            return []

    groups = load_groups()

    task_items = [
        ("任务1（必知必会）", "config/config.yaml"),
        ("任务2（DHIA Software）", "task_software/config/config.yaml"),
        ("任务3（客户培训覆盖）", "task_customer/config/config.yaml"),
    ]

    root = tk.Tk()
    root.title("培训分析一键运行")
    root.geometry("980x720")

    def apply_macos_like_style():
        style = ttk.Style()
        themes = set(style.theme_names() or [])
        if "clam" in themes:
            style.theme_use("clam")
        elif "vista" in themes:
            style.theme_use("vista")

        bg = "#F5F5F7"
        panel_bg = "#FFFFFF"
        border = "#D1D1D6"
        text = "#1C1C1E"
        subtle = "#3A3A3C"
        accent = "#0A84FF"
        accent_hover = "#409CFF"
        danger = "#FF3B30"
        danger_hover = "#FF6B63"

        root.configure(bg=bg)
        try:
            root.option_add("*Font", "{Segoe UI} 10")
        except Exception:
            pass

        try:
            style.configure(".", background=bg, foreground=text)
            style.configure("TFrame", background=bg)
            style.configure("TLabel", background=bg, foreground=text)
        except Exception:
            pass

        try:
            style.configure(
                "TLabelframe",
                background=bg,
                foreground=text,
                bordercolor=border,
                relief="solid",
            )
            style.configure("TLabelframe.Label", background=bg, foreground=subtle)
        except Exception:
            pass

        try:
            style.configure(
                "TButton",
                padding=(14, 8),
                background=panel_bg,
                foreground=text,
                bordercolor=border,
                focusthickness=2,
                focuscolor=accent,
            )
            style.map(
                "TButton",
                background=[("active", "#EAEAED"), ("disabled", "#E5E5EA")],
                foreground=[("disabled", "#8E8E93")],
            )
        except Exception:
            pass

        try:
            style.configure(
                "Accent.TButton",
                padding=(14, 8),
                background=accent,
                foreground="#FFFFFF",
                bordercolor=accent,
                focusthickness=2,
                focuscolor=accent,
            )
            style.map(
                "Accent.TButton",
                background=[("active", accent_hover), ("disabled", "#A7CFFF")],
                foreground=[("disabled", "#FFFFFF")],
            )
        except Exception:
            pass

        try:
            style.configure(
                "Danger.TButton",
                padding=(14, 8),
                background=danger,
                foreground="#FFFFFF",
                bordercolor=danger,
                focusthickness=2,
                focuscolor=danger,
            )
            style.map(
                "Danger.TButton",
                background=[("active", danger_hover), ("disabled", "#F2A6A1")],
                foreground=[("disabled", "#FFFFFF")],
            )
        except Exception:
            pass

        try:
            style.configure(
                "TRadiobutton",
                background=bg,
                foreground=text,
                padding=(6, 4),
            )
        except Exception:
            pass

        try:
            style.configure(
                "TCombobox",
                padding=(8, 6),
                foreground=text,
                fieldbackground=panel_bg,
                background=panel_bg,
                bordercolor=border,
                lightcolor=border,
                darkcolor=border,
            )
        except Exception:
            pass

        try:
            style.configure(
                "TEntry",
                padding=(8, 6),
                foreground=text,
                fieldbackground=panel_bg,
                background=panel_bg,
                bordercolor=border,
            )
        except Exception:
            pass

    apply_macos_like_style()
    try:
        root.update_idletasks()
        root.lift()
        root.focus_force()
        root.attributes("-topmost", True)
        root.update()
        root.attributes("-topmost", False)
    except Exception:
        pass

    state = {
        "proc": None,
        "thread": None,
        "queue": queue.Queue(),
        "running": False,
    }

    task_var = tk.StringVar(value=task_items[0][0])
    group_var = tk.StringVar(value="全部（所有三级部门）")
    exclude_var = tk.StringVar(value="无")
    email_mode_var = tk.StringVar(value="正式发送")

    def _is_all_group_label(v):
        return str(v or "").strip() in {"全部（所有三级部门）", "全部（所有大区）"}

    def build_command():
        config_path = dict(task_items).get(task_var.get(), "config/config.yaml")
        args = [sys.executable, "-X", "utf8", os.path.join(repo_root, "main.py")]
        if group_var.get() and not _is_all_group_label(group_var.get()):
            args.append(f"--小组={group_var.get()}")
        if exclude_var.get() and exclude_var.get() != "无":
            args.append(f"--排除小组={exclude_var.get()}")
        mode = email_mode_var.get()
        if mode == "测试邮件（仅Morgan）":
            args.append("--test-email")
        elif mode == "不发邮件":
            args.append("--no-email")
        elif mode == "正式发送" and config_path.replace("\\", "/").endswith("task_customer/config/config.yaml"):
            args.append("--force-email")
        args.append(f"--config={config_path}")
        return args

    def refresh_preview(*_):
        cmd = build_command()
        preview_var.set(" ".join([f'"{c}"' if " " in c else c for c in cmd]))

    def set_running(running: bool):
        state["running"] = running
        run_btn.config(state=("disabled" if running else "normal"))
        stop_btn.config(state=("normal" if running else "disabled"))
        task_cb.config(state=("disabled" if running else "readonly"))
        group_cb.config(state=("disabled" if running else "readonly"))
        exclude_cb.config(state=("disabled" if running else "readonly"))
        for rb in email_rbs:
            rb.config(state=("disabled" if running else "normal"))

    def reader_thread(proc):
        try:
            for line in iter(proc.stdout.readline, ""):
                if not line:
                    break
                state["queue"].put(line)
        finally:
            try:
                proc.stdout.close()
            except Exception:
                pass
            state["queue"].put(None)

    def pump_queue():
        try:
            while True:
                item = state["queue"].get_nowait()
                if item is None:
                    set_running(False)
                    state["proc"] = None
                    return
                log_text.configure(state="normal")
                log_text.insert("end", item)
                log_text.see("end")
                log_text.configure(state="disabled")
        except queue.Empty:
            pass
        root.after(60, pump_queue)

    def on_run():
        if state["running"]:
            return
        cmd = build_command()
        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.insert("end", "执行命令：\n" + " ".join(cmd) + "\n\n")
        log_text.configure(state="disabled")

        try:
            env = os.environ.copy()
            env["PYTHONUTF8"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"
            proc = subprocess.Popen(
                cmd,
                cwd=repo_root,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
        except Exception as e:
            messagebox.showerror("启动失败", str(e))
            return

        state["proc"] = proc
        set_running(True)
        t = threading.Thread(target=reader_thread, args=(proc,), daemon=True)
        state["thread"] = t
        t.start()
        root.after(60, pump_queue)

    def on_stop():
        proc = state.get("proc")
        if not proc:
            return
        try:
            proc.terminate()
        except Exception:
            pass

    def open_output_dir():
        config_path = dict(task_items).get(task_var.get(), "config/config.yaml")
        try:
            with open(os.path.join(repo_root, config_path), "r", encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            out_dir = (cfg.get("paths", {}) or {}).get("output_dir", "outputs/")
            out_dir = os.path.abspath(os.path.join(repo_root, out_dir))
            os.startfile(out_dir)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    container = ttk.Frame(root, padding=0)
    container.pack(fill="both", expand=True)

    content = ttk.Frame(container, padding=12)
    content.pack(fill="both", expand=True)

    def draw_horizontal_gradient(canvas, width, height, left_color, right_color):
        def hex_to_rgb(h):
            h = h.lstrip("#")
            return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

        def rgb_to_hex(r, g, b):
            return f"#{r:02x}{g:02x}{b:02x}"

        lr, lg, lb = hex_to_rgb(left_color)
        rr, rg, rb = hex_to_rgb(right_color)
        steps = max(1, width)
        for x in range(steps):
            t = x / (steps - 1) if steps > 1 else 0
            r = int(lr + (rr - lr) * t)
            g = int(lg + (rg - lg) * t)
            b = int(lb + (rb - lb) * t)
            canvas.create_line(x, 0, x, height, fill=rgb_to_hex(r, g, b))

    header = tk.Canvas(container, height=92, highlightthickness=0, bd=0)
    header.pack(fill="x")

    def render_header(event=None):
        header.delete("all")
        w = max(1, header.winfo_width())
        h = 92
        draw_horizontal_gradient(header, w, h, "#0A84FF", "#BF5AF2")
        header.create_text(18, 26, anchor="w", text="培训分析一键运行", fill="#FFFFFF", font=("{Segoe UI}", 18, "bold"))
        header.create_text(18, 56, anchor="w", text="选择任务与范围，一键生成汇总/PPT并发送邮件", fill="#F2F2F7", font=("{Segoe UI}", 10))

    header.bind("<Configure>", render_header)

    top = ttk.Frame(content)
    top.pack(fill="x")

    ttk.Label(top, text="任务：").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
    task_cb = ttk.Combobox(top, textvariable=task_var, values=[t[0] for t in task_items], state="readonly", width=28)
    task_cb.grid(row=0, column=1, sticky="w", pady=6)

    group_label = ttk.Label(top, text="小组（三级部门）：")
    group_label.grid(row=0, column=2, sticky="w", padx=(18, 8), pady=6)
    group_values = ["全部（所有三级部门）"] + groups
    group_cb = ttk.Combobox(top, textvariable=group_var, values=group_values, state="readonly", width=28)
    group_cb.grid(row=0, column=3, sticky="w", pady=6)

    ttk.Label(top, text="排除小组：").grid(row=0, column=4, sticky="w", padx=(18, 8), pady=6)
    exclude_values = ["无"] + groups
    exclude_cb = ttk.Combobox(top, textvariable=exclude_var, values=exclude_values, state="readonly", width=18)
    exclude_cb.grid(row=0, column=5, sticky="w", pady=6)

    def _load_customer_regions():
        config_path = dict(task_items).get(task_var.get(), "config/config.yaml")
        try:
            with open(os.path.join(repo_root, config_path), "r", encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            if cfg.get("task_type") != "customer_coverage":
                return None
            owners = (cfg.get("email", {}) or {}).get("group_owners", {}) or {}
            regions = [str(k).strip() for k in owners.keys() if str(k).strip() and str(k).strip() != "总部"]
            regions = sorted(set(regions))
            return regions
        except Exception:
            return None

    def _on_task_changed(*_):
        config_path = dict(task_items).get(task_var.get(), "config/config.yaml")
        regions = _load_customer_regions()
        if regions is not None:
            group_label.config(text="大区：")
            all_label = "全部（所有大区）"
            new_group_values = [all_label] + regions
            new_exclude_values = ["无"] + regions
        else:
            group_label.config(text="小组（三级部门）：")
            all_label = "全部（所有三级部门）"
            new_group_values = [all_label] + groups
            new_exclude_values = ["无"] + groups

        group_cb.config(values=new_group_values)
        exclude_cb.config(values=new_exclude_values)
        if group_var.get() not in new_group_values:
            group_var.set(all_label)
        if exclude_var.get() not in new_exclude_values:
            exclude_var.set("无")
        refresh_preview()

    task_var.trace_add("write", _on_task_changed)

    email_frame = ttk.Labelframe(content, text="邮件选项", padding=10)
    email_frame.pack(fill="x", pady=(10, 8))
    email_rbs = []
    for i, label in enumerate(["正式发送", "测试邮件（仅Morgan）", "不发邮件"]):
        rb = ttk.Radiobutton(email_frame, text=label, variable=email_mode_var, value=label, command=refresh_preview)
        rb.grid(row=0, column=i, sticky="w", padx=(0 if i == 0 else 18, 0), pady=4)
        email_rbs.append(rb)

    preview_frame = ttk.Labelframe(content, text="命令预览", padding=10)
    preview_frame.pack(fill="x", pady=(0, 10))
    preview_var = tk.StringVar()
    preview_entry = ttk.Entry(preview_frame, textvariable=preview_var, state="readonly")
    preview_entry.pack(fill="x")

    btns = ttk.Frame(content)
    btns.pack(fill="x", pady=(0, 10))
    run_btn = ttk.Button(btns, text="执行", command=on_run, style="Accent.TButton")
    run_btn.pack(side="left")
    stop_btn = ttk.Button(btns, text="停止", command=on_stop, state="disabled", style="Danger.TButton")
    stop_btn.pack(side="left", padx=(10, 0))
    ttk.Button(btns, text="打开输出目录", command=open_output_dir).pack(side="left", padx=(10, 0))

    mid = ttk.Frame(content)
    mid.pack(fill="both", expand=True)

    photo_frame = ttk.Labelframe(mid, text="图片", padding=10)
    photo_frame.pack(side="right", fill="y", padx=(10, 0))

    photo_title = ttk.Label(photo_frame, text="可嵌入一张图片作为页面装饰")
    photo_title.pack(anchor="w", pady=(0, 8))

    try:
        from PIL import Image, ImageTk
    except Exception:
        Image = None
        ImageTk = None

    photo_label = ttk.Label(photo_frame)
    photo_label.pack(pady=(0, 10))

    photo_state = {"imgtk": None}

    def set_photo_from_path(path):
        if not path:
            photo_state["imgtk"] = None
            photo_label.configure(image="", text="未加载图片")
            return
        if Image is None or ImageTk is None:
            messagebox.showwarning("缺少依赖", "当前环境缺少 Pillow，无法加载 JPG/PNG 图片。")
            return
        try:
            img = Image.open(path).convert("RGB")
            img = img.resize((240, 240), Image.LANCZOS)
            imgtk = ImageTk.PhotoImage(img)
            photo_state["imgtk"] = imgtk
            photo_label.configure(image=imgtk, text="")
        except Exception as e:
            messagebox.showerror("加载失败", str(e))

    def on_choose_photo():
        path = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[
                ("Image files", "*.png;*.jpg;*.jpeg;*.gif;*.webp;*.bmp"),
                ("All files", "*.*"),
            ],
        )
        if path:
            set_photo_from_path(path)

    if Image is None or ImageTk is None:
        photo_label.configure(text="未检测到 Pillow，无法加载图片")
    else:
        photo_label.configure(text="未加载图片")

    ttk.Button(photo_frame, text="选择图片…", command=on_choose_photo).pack(fill="x")

    note = ttk.Label(
        photo_frame,
        text="提示：请使用你有权使用的图片文件。\n（例如公司内网素材、自拍或授权图片）",
        foreground="#3A3A3C",
        justify="left",
    )
    note.pack(anchor="w", pady=(10, 0))

    log_frame = ttk.Labelframe(mid, text="运行日志", padding=10)
    log_frame.pack(side="left", fill="both", expand=True)
    log_text = ScrolledText(log_frame, wrap="word")
    log_text.pack(fill="both", expand=True)
    try:
        log_text.configure(
            background="#FFFFFF",
            foreground="#1C1C1E",
            insertbackground="#1C1C1E",
            padx=10,
            pady=10,
            font=("Consolas", 10),
        )
    except Exception:
        pass
    log_text.configure(state="disabled")

    status_bar = ttk.Frame(content)
    status_bar.pack(fill="x", pady=(8, 0))
    status_label = ttk.Label(status_bar, text="就绪", foreground="#3A3A3C")
    status_label.pack(side="left")

    task_var.trace_add("write", refresh_preview)
    group_var.trace_add("write", refresh_preview)
    exclude_var.trace_add("write", refresh_preview)

    refresh_preview()

    def on_close():
        proc = state.get("proc")
        if proc and proc.poll() is None:
            if not messagebox.askyesno("退出", "任务仍在运行，确定要退出吗？"):
                return
            try:
                proc.terminate()
            except Exception:
                pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

def main():
    _configure_stdio()
    """支持多数据源汇总、三级部门解析及自动分析说明的工具"""
    # 检查是否是测试模式
    test_mode = "--test" in sys.argv
    rep_mode = "--代表处" in sys.argv
    email_disabled = "--no-email" in sys.argv
    test_email_only = "--test-email" in sys.argv
    email_preview = "--email-preview" in sys.argv
    force_email = "--force-email" in sys.argv
    test_recipient = "morgan.hu@dahuatech.com"
    group_scope = None
    exclude_scope = None

    rep_scope = None
    for arg in sys.argv:
        if arg.startswith("--代表处范围="):
            raw = arg.split("=", 1)[1]
            raw = raw.replace("，", ",").replace(";", ",").replace("；", ",")
            parts = [p.strip() for p in raw.split(",") if p.strip()]
            if parts:
                rep_scope = rep_scope or set()
                rep_scope.update(parts)
        if arg.startswith("--排除小组范围=") or arg.startswith("--排除大区范围="):
            raw = arg.split("=", 1)[1]
            raw = raw.replace("，", ",").replace(";", ",").replace("；", ",")
            parts = [p.strip() for p in raw.split(",") if p.strip()]
            if parts:
                exclude_scope = exclude_scope or set()
                exclude_scope.update(parts)
        if arg.startswith("--排除小组="):
            raw = arg.split("=", 1)[1].strip()
            if raw:
                exclude_scope = exclude_scope or set()
                exclude_scope.add(raw)
        if arg.startswith("--小组范围=") or arg.startswith("--大区范围="):
            raw = arg.split("=", 1)[1]
            raw = raw.replace("，", ",").replace(";", ",").replace("；", ",")
            parts = [p.strip() for p in raw.split(",") if p.strip()]
            if parts:
                group_scope = group_scope or set()
                group_scope.update(parts)
        if arg.startswith("--小组="):
            raw = arg.split("=", 1)[1].strip()
            if raw:
                group_scope = group_scope or set()
                group_scope.add(raw)
        if arg.startswith("--test-recipient="):
            raw = arg.split("=", 1)[1].strip()
            if raw:
                test_recipient = raw
    
    # 允许通过命令行指定配置文件，默认为 config/config.yaml
    config_path = "config/config.yaml"
    for arg in sys.argv:
        if arg.startswith("--config="):
            config_path = arg.split("=")[1]
    
    if not os.path.exists(config_path):
        print(f"❌ 错误: 找不到配置文件 '{config_path}'")
        return

    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    task_name = config.get('task_name', '培训数据分析汇总工具 v3.0')
    os.makedirs(config['paths']['output_dir'], exist_ok=True)
    os.makedirs(config['paths']['data_dir'], exist_ok=True)

    print("=" * 60)
    if test_mode:
        print(f"🚀 [测试模式] {task_name} - 正在执行格式确认测试")
    else:
        print(f"{task_name}")
    print("=" * 60)

    if config.get("task_type") == "customer_coverage":
        try:
            from modules.customer_coverage_analyzer import CustomerCoverageAnalyzer
            from modules.customer_exporter import CustomerExporter
            from reports.customer_ppt_generator import CustomerPPTGenerator

            analysis_cfg = config.get("analysis", {}) or {}
            enable_incremental = analysis_cfg.get("cert_mapping_incremental", True)
            if enable_incremental:
                try:
                    from task_customer.build_cert_mapping_incremental import run_incremental

                    print("\n[阶段0] 正在执行证书增量映射（按GSP ID）...")
                    run_incremental(config_path)
                except Exception as e:
                    print(f"⚠️ 证书增量映射执行失败，将继续使用现有映射文件: {e}")

            print("\n[阶段1] 正在分析客户培训覆盖与认证情况...")
            analyzer = CustomerCoverageAnalyzer(config_path)
            dfs = analyzer.analyze()
            print("✅ 数据分析完成")

            print("\n[阶段2] 正在导出 Excel 汇总与明细...")
            exporter = CustomerExporter(config_path)
            out_files = exporter.export_excels(dfs)
            print(f"✅ 汇总Excel: {out_files.get('summary_excel')}")
            print(f"✅ 明细Excel: {out_files.get('detail_excel')}")

            print("\n[阶段3] 正在生成 PPT 汇报...")
            ppt_gen = CustomerPPTGenerator(config_path)
            out_dir = config.get("paths", {}).get("output_dir", "outputs/")
            output_ppt = os.path.join(out_dir, "customer_coverage_report.pptx")
            ppt_gen.generate_report(
                dfs,
                template_path=config.get("paths", {}).get("template_ppt"),
                output_path=output_ppt,
            )
            print(f"✅ PPT: {output_ppt}")

            print("\n[阶段4] 正在准备发送邮件通知...")
            if email_disabled:
                print("ℹ️ 已指定 --no-email，跳过邮件发送。")
            else:
                try:
                    from modules.customer_email_sender import CustomerEmailSender

                    email_sender = CustomerEmailSender(config_path)
                    region_filter = None
                    if group_scope:
                        region_filter = set(group_scope)
                    elif test_email_only:
                        region_filter = {"非洲区"}
                    email_sender.send_customer_coverage_notifications(
                        dfs,
                        region_filter=region_filter,
                        test_email_only=test_email_only,
                        test_recipient=test_recipient,
                        preview_only=email_preview,
                        force_send=force_email,
                    )
                except Exception as e:
                    print(f"❌ 邮件发送流程出错: {e}")

            print("\n" + "=" * 60)
            print("分析任务圆满完成！")
            print("=" * 60)
            return
        except Exception as e:
            print(f"❌ 客户培训覆盖任务执行失败: {e}")
            import traceback
            traceback.print_exc()
            return

    analyzer = DataAnalyzer(config_path)
    all_results = []
    all_raw_dfs = []
    
    analysis_filter = None
    notify_filter = None
    if rep_mode and rep_scope:
        notify_filter = set(rep_scope)
    elif group_scope:
        notify_filter = set(group_scope)
    
    scope_tag = None
    if notify_filter:
        scope_tag = "_".join(sorted(notify_filter))
    elif exclude_scope:
        scope_tag = "排除" + "_".join(sorted(exclude_scope))
        for ch in ['\\', '/', '?', '*', '[', ']', ':', '"', '<', '>', '|']:
            scope_tag = scope_tag.replace(ch, ' ')
        scope_tag = scope_tag.strip().replace(' ', '')


    # 1. 循环处理所有输入文件 (data1.xlsx, data2.xlsx)
    input_files = config['paths'].get('input_files', [])
    for file_path in input_files:
        if not os.path.exists(file_path):
            print(f"⚠️ 跳过: 文件不存在 '{file_path}'")
            continue
        
        print(f"\n[阶段1] 正在处理原始数据: {file_path}")
        try:
            # 该步骤会解析三级部门并统计每个文件的基础数据
            result_df, raw_df = analyzer.analyze_learning_progress(file_path, group_filter=analysis_filter)
            all_results.append(result_df)
            all_raw_dfs.append(raw_df)
            print(f"✅ 已提取 {len(result_df)} 个部门的数据")
        except Exception as e:
            print(f"❌ 处理文件 {file_path} 出错: {e}")

    if not all_results:
        print("\n❌ 错误: 没有成功处理任何数据文件，请检查 data 文件夹。")
        return

    # 1.5 导出各小组未完成学员清单
    print("\n[阶段1.5] 正在生成各小组未完成学员清单...")
    try:
        saved_files = analyzer.save_uncompleted_lists(all_raw_dfs, config['paths']['output_dir'], group_filter=notify_filter, exclude_filter=exclude_scope)
        print(f"✅ 已生成 {len(saved_files)} 份小组未完成学员清单，保存在 outputs/uncompleted_lists/")
    except Exception as e:
        print(f"❌ 生成未完成清单出错: {e}")

    # 2. 汇总所有数据并排名
    print("\n[阶段2] 正在汇总数据并对比上一轮排名...")
    try:
        # 尝试加载上一轮数据
        previous_df = None
        prev_path = config['paths'].get('previous_analysis')
        if prev_path and os.path.exists(prev_path):
            try:
                previous_df = pd.read_excel(prev_path)
                print(f"ℹ️ 已加载上一轮汇总数据: {prev_path}")
            except Exception as e:
                print(f"⚠️ 加载上一轮数据失败: {e}")

        final_df = analyzer.merge_results(all_results, previous_df=previous_df)

        extra_sheets = None
        if rep_mode:
            combined_raw = pd.concat(all_raw_dfs, ignore_index=True)
            status_col = analyzer.detect_status_col(combined_raw)

            if '三级部门' not in combined_raw.columns or '四级部门' not in combined_raw.columns:
                raise ValueError("原始数据缺少“三级部门/四级部门”字段，无法生成代表处明细 sheet")

            extra_sheets = {}
            rep_sheet_scope = set(rep_scope) if rep_scope else None
            level3_names = [g for g in final_df['小组'].tolist() if g != '总计']
            for level3 in level3_names:
                if rep_sheet_scope and level3 not in rep_sheet_scope:
                    continue
                sub_df = combined_raw[combined_raw['三级部门'] == level3]
                if sub_df.empty:
                    continue

                base_level4 = analyzer.build_group_summary(sub_df, '四级部门', status_col)

                prev_sheet_df = None
                if prev_path and os.path.exists(prev_path):
                    try:
                        prev_sheet_df = pd.read_excel(prev_path, sheet_name=analyzer.sanitize_sheet_name(level3))
                    except Exception:
                        prev_sheet_df = None

                level4_final = analyzer.merge_results([base_level4], previous_df=prev_sheet_df)
                extra_sheets[level3] = level4_final

        excel_base = "combined_analysis"
        if scope_tag:
            excel_base = f"{excel_base}_{scope_tag}"
        output_excel_path = os.path.join(config['paths']['output_dir'], f"{excel_base}.xlsx")
        try:
            analyzer.save_analysis_excel(final_df, output_excel_path, extra_sheets=extra_sheets)
        except PermissionError:
            import time
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_excel_path = os.path.join(config['paths']['output_dir'], f"{excel_base}_{timestamp}.xlsx")
            analyzer.save_analysis_excel(final_df, output_excel_path, extra_sheets=extra_sheets)
            print(f"⚠️ 原始文件被占用，已另存为: {output_excel_path}")

        print(f"✅ 汇总完成！已按本轮完成率排名。")
        print(f"✅ 汇总结果已保存至: {output_excel_path}")
    except Exception as e:
        print(f"❌ 汇总数据出错: {e}")
        import traceback
        traceback.print_exc()
        return

    # 3. 生成 PPT 汇报
    print("\n[阶段3] 正在生成 PPT 汇报 (含自动分析说明)...")
    try:
        ppt_gen = PPTGenerator(config_path)
        ppt_base = "final_weekly_report"
        if scope_tag:
            ppt_base = f"{ppt_base}_{scope_tag}"
        output_ppt_path = os.path.join(config['paths']['output_dir'], f"{ppt_base}.pptx")
        
        try:
            # 使用 data/PPT模板.pptx 作为模板
            ppt_gen.generate_report(final_df, template_path="data/PPT模板.pptx", output_path=output_ppt_path)
        except PermissionError:
            import time
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_ppt_path = os.path.join(config['paths']['output_dir'], f"{ppt_base}_{timestamp}.pptx")
            ppt_gen.generate_report(final_df, template_path="data/PPT模板.pptx", output_path=output_ppt_path)
            print(f"⚠️ 原始 PPT 文件被占用，已另存为: {output_ppt_path}")

        print(f"✅ PPT 汇报已生成至: {output_ppt_path}")
        print("💡 提示：请打开 PPT 查看右侧自动生成的分析文字。")
    except Exception as e:
        print(f"❌ PPT 生成出错: {e}")
        import traceback
        traceback.print_exc()
        return

    # 4. 发送邮件通知 (可选功能)
    print("\n[阶段4] 正在准备发送邮件通知...")
    if email_disabled:
        print("ℹ️ 已指定 --no-email，跳过邮件发送。")
        print("\n" + "=" * 60)
        print("分析任务圆满完成！")
        print("=" * 60)
        return
    try:
        from modules.email_sender import EmailSender
        email_sender = EmailSender(config_path)
        uncompleted_dir = os.path.join(config['paths']['output_dir'], "uncompleted_lists")
        # 排除总计行，发送各小组通知
        email_sender.send_all_notifications(
            final_df,
            uncompleted_dir,
            all_raw_dfs,
            test_mode=test_mode,
            rep_mode=rep_mode,
            group_filter=notify_filter,
            exclude_filter=exclude_scope,
            test_email_only=test_email_only,
            test_recipient=test_recipient
        )
    except Exception as e:
        print(f"❌ 邮件发送流程出错: {e}")

    print("\n" + "=" * 60)
    print("分析任务圆满完成！")
    print("=" * 60)

if __name__ == "__main__":
    if "--gui" in sys.argv:
        launch_gui()
    else:
        main()
