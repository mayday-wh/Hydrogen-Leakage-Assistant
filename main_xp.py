# -*- coding: utf-8 -*-
"""
软件名称：发电机漏氢统计助手
版本号：v1.0
功能描述：实现电厂发电机组漏氢数据的实时记录、自动计算、月/年汇总分析及报表导出。
"""

import sys
import os
import json
import xlsxwriter
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
import ctypes

# --- 系统环境配置：解决不同分辨率（DPI）屏幕下的界面问题 ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1) 
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

def get_base_dir():
    """ 
    动态获取程序运行根目录，兼容脚本运行和打包后的exe环境。
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

# 目录与索引文件定义
BASE_DIR = get_base_dir()
DATA_DIR = os.path.join(BASE_DIR, "system_data")    # 存储各机组历史 JSON 数据
CONFIG_DIR = os.path.join(BASE_DIR, "config")      # 存储机组容积等配置参数
INDEX_FILE = os.path.join(CONFIG_DIR, "units_index.json") # 机组名称索引

# 初始化系统目录
for d in [DATA_DIR, CONFIG_DIR]:
    if not os.path.exists(d): os.makedirs(d)

class HydrogenAppXP:
    """ 
    主程序逻辑类：负责UI渲染、逻辑计算及数据交互
    """
    def __init__(self, root):
        self.root = root
        self.root.title("发电机漏氢统计助手 v1.0")
        
        # 屏幕适配：根据当前分辨率自动调整全局缩放比例
        screen_w = self.root.winfo_screenwidth()
        self.scale = max(0.6, min(1.0, screen_w / 3840))
        
        # 设置窗口初始大小及最小尺寸限制
        win_w = int(1380 * self.scale)
        win_h = int(1160 * self.scale)
        self.root.geometry("{0}x{1}".format(win_w, win_h))
        self.root.minsize(int(1000 * self.scale), int(700 * self.scale))
        
        # 统一定义字体样式
        self.font_title = ("微软雅黑", 22, "bold")
        self.font_main = ("微软雅黑", 12, "bold")
        self.font_normal = ("微软雅黑", 11)
        self.font_val = ("微软雅黑", 12, "bold")
        
        # 状态变量初始化
        self.is_admin = False               # 权限标识：False为班组模式，True为专工模式
        self.admin_pwd = "chd123456"        # 默认管理员密码
        self.units = self.load_units()      # 加载机组列表
        self.current_unit = ""              # 当当前选中的机组名
        self.data_list = []                 # 存储当前机组的所有历史记录
        self.current_idx = -1               # 当前显示的记录索引
        self.unit_btns = {}                 # 侧边栏按钮对象池
        
        # 默认机组参数
        self.params = {"volume": 125.0, "p_atm": 0.101325, "base_temp": 40.0}
        
        # 预设班组列表
        self.teams = ["燃机一班", "燃机二班", "燃机三班", "燃机四班", "燃机五班",
                      "煤机一班", "煤机二班", "煤机三班", "煤机四班", "煤机五班"]
        
        # 界面显示变量绑定（用于Tkinter自动更新）
        self.data_vars = {}
        keys = ["t1", "t2", "d1", "d2", "p1", "p2", "temp1", "temp2", "h", "fill", "base", "comp", "team",
                "avg_m_base", "avg_m_comp", "avg_y_base", "avg_y_comp"]
        for k in keys:
            self.data_vars[k] = tk.StringVar(value="--")
        
        self.setup_ui()                     # 初始化界面
        if self.units: self.switch_unit(self.units[0])  # 默认加载首个机组
        self.refresh_ui_state()             # 更新按钮

    def setup_ui(self):
        """ 构造主界面布局 """
        # --- 左侧：机组导航栏 ---
        sidebar_w = int(220 * self.scale)
        self.side_bar = tk.Frame(self.root, width=sidebar_w, bg="#2C3E50")
        self.side_bar.pack(side="left", fill="y")
        self.side_bar.pack_propagate(False)
        tk.Label(self.side_bar, text="机组导航", bg="#2C3E50", fg="white", font=self.font_main).pack(pady=20)
        
        self.unit_scroll_fm = tk.Frame(self.side_bar, bg="#2C3E50")
        self.unit_scroll_fm.pack(fill="both", expand=True)
        self.refresh_unit_sidebar()
        
        # 管理功能按钮（需专工权限）
        self.btn_add_unit = tk.Button(self.side_bar, text="添加机组", command=self.add_unit_ui, font=self.font_normal)
        self.btn_add_unit.pack(fill="x", padx=10, pady=2)
        self.btn_rename_unit = tk.Button(self.side_bar, text="修改名称", command=self.rename_unit_ui, font=self.font_normal)
        self.btn_rename_unit.pack(fill="x", padx=10, pady=2)
        
        # 权限切换及模式显示
        tk.Button(self.side_bar, text="切换权限", command=self.toggle_admin, bg="#34495E", fg="white", font=self.font_normal).pack(fill="x", padx=10, pady=10)
        self.admin_label = tk.Label(self.side_bar, text="[班组模式]", bg="#2C3E50", fg="#95A5A6", font=self.font_normal)
        self.admin_label.pack(side="bottom", pady=10)

        # --- 右侧：数据展示与操作区 ---
        self.right_main_container = tk.Frame(self.root)
        self.right_main_container.pack(side="right", fill="both", expand=True)

        # 底部操作栏：包含翻页、增删改查及导出功能
        ctl_area = ttk.Frame(self.right_main_container, padding=10)
        ctl_area.pack(side="bottom", fill="x", pady=5)
        
        # 记录导航（上条/下条）
        nav_fm = tk.Frame(ctl_area)
        nav_fm.pack(fill="x", pady=5)
        self.page_info = tk.Label(nav_fm, text="0 / 0", font=self.font_main)
        self.page_info.pack(side="left", padx=10)
        
        btn_w = int(14 * self.scale + 4)
        tk.Button(nav_fm, text="下条记录", width=btn_w, font=self.font_normal, command=lambda: self.move_idx(1)).pack(side="right", padx=5)
        tk.Button(nav_fm, text="上条记录", width=btn_w, font=self.font_normal, command=lambda: self.move_idx(-1)).pack(side="right", padx=5)
        
        # 业务操作按钮
        ops_fm = tk.Frame(ctl_area)
        ops_fm.pack(fill="x", pady=5)
        
        self.btn_cfg = tk.Button(ops_fm, text="机组参数", bg="#7F8C8D", fg="white", font=self.font_normal, command=self.open_config_ui)
        self.btn_cfg.pack(side="left", padx=5)

        btns_right_container = tk.Frame(ops_fm)
        btns_right_container.pack(side="right")

        # 功能指令按钮
        btns_info = [
            ("删除当前", "#C0392B", self.delete_record, "btn_del"),
            ("修改当前", "#D35400", self.open_edit_window, None),
            ("新增记录", "#27AE60", self.open_record_window, None),
            ("导出报表", "#2980B9", self.open_export_dialog, None)
        ]

        for text, color, cmd, attr_name in btns_info:
            btn = tk.Button(btns_right_container, text=text, bg=color, fg="white", font=self.font_normal, padx=10, command=cmd)
            btn.pack(side="right", padx=3)
            if attr_name:
                setattr(self, attr_name, btn)

        # 中央数据详情区：带滚动条
        self.canvas = tk.Canvas(self.right_main_container, highlightthickness=0)
        self.v_scrollbar = ttk.Scrollbar(self.right_main_container, orient="vertical", command=self.canvas.yview)
        pad_size = int(30 * self.scale)
        self.main_fm = ttk.Frame(self.canvas, padding=pad_size)
        self.v_scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.main_fm, anchor="nw")
        
        # 动态绑定画布滚动逻辑
        self.main_fm.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.canvas_frame, width=e.width))

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 数据面板：显示机组名称及具体参数
        self.title_var = tk.StringVar(value="请选择机组")
        tk.Label(self.main_fm, textvariable=self.title_var, font=self.font_title, fg="#2980B9").pack(anchor="w")

        # 参数展示面板 1：补氢前后原始数据
        dp_top = ttk.LabelFrame(self.main_fm, text=" 补氢前后参数 ", padding=10)
        dp_top.pack(fill="x", pady=10)
        for i in range(3): dp_top.columnconfigure(i, weight=1)
        
        tk.Label(dp_top, text="项目", font=self.font_main).grid(row=0, column=0, sticky="w")
        tk.Label(dp_top, text="补氢前", font=self.font_main, fg="#D35400").grid(row=0, column=1, sticky="w")
        tk.Label(dp_top, text="补氢后", font=self.font_main, fg="#2980B9").grid(row=0, column=2, sticky="w")
        
        rows = [("氢压 (MPa):", "p1", "p2"), ("氢温 (℃):", "temp1", "temp2"), ("记录日期:", "d1", "d2"), ("记录时间:", "t1", "t2")]
        for i, (label, v1, v2) in enumerate(rows):
            tk.Label(dp_top, text=label, font=self.font_normal).grid(row=i+1, column=0, pady=5, sticky="w")
            tk.Label(dp_top, textvariable=self.data_vars[v1], font=self.font_val).grid(row=i+1, column=1, sticky="w")
            tk.Label(dp_top, textvariable=self.data_vars[v2], font=self.font_val).grid(row=i+1, column=2, sticky="w")

        # 参数展示面板 2：漏氢统计结果
        dp_res = ttk.LabelFrame(self.main_fm, text=" 漏氢统计结果 ", padding=10)
        dp_res.pack(fill="x", pady=10)
        dp_res.columnconfigure(0, weight=1); dp_res.columnconfigure(1, weight=1)
        self.calc_labels = {}
        res_layout = [("漏氢时长 (h):", "h", 0, 0), ("本次补氢量 (m³):", "fill", 0, 1), 
                      ("基准率 (m³/d):", "base", 1, 0), ("补偿率 (m³/d):", "comp", 1, 1), 
                      ("月均基准:", "avg_m_base", 2, 0), ("月均补偿:", "avg_m_comp", 2, 1), 
                      ("年均基准:", "avg_y_base", 3, 0), ("年均补偿:", "avg_y_comp", 3, 1)]
        
        for label, v, r, c in res_layout:
            cell_fm = tk.Frame(dp_res)
            cell_fm.grid(row=r, column=c, pady=5, sticky="w")
            tk.Label(cell_fm, text=label, font=self.font_normal).pack(side="left")
            val_lb = tk.Label(cell_fm, textvariable=self.data_vars[v], font=self.font_val, fg="#27AE60")
            val_lb.pack(side="left", padx=5); self.calc_labels[v] = val_lb

        # 班组信息显示
        self.team_info_fm = tk.Frame(self.main_fm)
        self.team_info_fm.pack(anchor="e", padx=10, pady=5)
        tk.Label(self.team_info_fm, text="补氢班组:", font=self.font_normal).pack(side="left")
        tk.Label(self.team_info_fm, textvariable=self.data_vars['team'], font=self.font_val, fg="#8E44AD").pack(side="left", padx=5)

    def _create_record_popup(self, title, old_data=None):
        """ 弹出式窗口：用于新增或修改漏氢数据记录 """
        win = tk.Toplevel(self.root)
        win.title(title)
        p_w, p_h = int(1050 * self.scale), int(620 * self.scale)
        win.geometry("{0}x{1}".format(p_w, p_h))
        
        # 录入提示
        tip_fm = tk.Frame(win, bg="#FFF9C4", pady=int(10*self.scale))
        tip_fm.pack(fill="x")
        tk.Label(tip_fm, text="提示：氢温请统一使用 DCS [发电机热风区温度 #33] 测点数值", bg="#FFF9C4", fg="#5D4037", font=self.font_main).pack()
        
        # 班组选择
        tk.Label(win, text="补氢班组:", font=self.font_normal).pack(pady=int(10*self.scale))
        team_cb = ttk.Combobox(win, values=self.teams, width=30, font=self.font_normal)
        team_cb.set(old_data.get('Team', self.teams[0]) if old_data else self.teams[0])
        team_cb.pack()

        # 数据输入网格
        grid_fm = tk.Frame(win, pady=int(10*self.scale))
        grid_fm.pack(fill="x", padx=int(80*self.scale))
        for i in range(3): grid_fm.columnconfigure(i, weight=1)
        
        tk.Label(grid_fm, text="项目", font=self.font_main).grid(row=0, column=0, sticky="w")
        tk.Label(grid_fm, text="补氢前", font=self.font_main, fg="#D35400").grid(row=0, column=1, sticky="w")
        tk.Label(grid_fm, text="补氢后", font=self.font_main, fg="#2980B9").grid(row=0, column=2, sticky="w")
        
        rows_cfg = [("氢压 (MPa):", "P1", "P2"), ("氢温 (℃):", "Temp1", "Temp2"), ("日期:", "D1", "D2"), ("时间:", "T1", "T2")]
        entries = {}
        cd, ct = datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M")
        for i, (label, k1, k2) in enumerate(rows_cfg):
            tk.Label(grid_fm, text=label, font=self.font_normal).grid(row=i+1, column=0, sticky="w", pady=5)
            e1, e2 = tk.Entry(grid_fm, width=12, font=self.font_normal), tk.Entry(grid_fm, width=12, font=self.font_normal)
            if old_data:
                e1.insert(0, str(old_data.get(k1,''))); e2.insert(0, str(old_data.get(k2,'')))
            else:
                if k1 in ["D1", "D2"]: e1.insert(0, cd); e2.insert(0, cd)
                elif k1 in ["T1", "T2"]: e1.insert(0, ct); e2.insert(0, ct)
            e1.grid(row=i+1, column=1, sticky="w"); e2.grid(row=i+1, column=2, sticky="w")
            entries[k1], entries[k2] = e1, e2

        def save():
            """ 数据有效性校验并保存 """
            try:
                nr = {"Team": team_cb.get(), "D1": entries['D1'].get(), "T1": entries['T1'].get(), "D2": entries['D2'].get(), "T2": entries['T2'].get(), "P1": float(entries['P1'].get()), "P2": float(entries['P2'].get()), "Temp1": float(entries['Temp1'].get()), "Temp2": float(entries['Temp2'].get())}
                if old_data: self.data_list[self.data_list.index(old_data)] = nr
                else: self.data_list.append(nr)
                self.recalc_and_save(); self.current_idx = len(self.data_list)-1; self.update_display(); win.destroy()
            except: messagebox.showerror("错误", "参数输入有误，请确保数值格式正确")
        
        tk.Button(win, text="确认保存记录", bg="#27AE60", fg="white", font=self.font_main, padx=30, pady=10, command=save).pack(pady=10)

    def refresh_ui_state(self):
        """ 切换管理按钮的锁定/解锁状态 """
        s = "normal" if self.is_admin else "disabled"
        for btn in [self.btn_del, self.btn_cfg, self.btn_add_unit, self.btn_rename_unit]:
            btn.config(state=s)
        self.admin_label.config(text="[专工模式]" if self.is_admin else "[班组模式]", fg="red" if self.is_admin else "#95A5A6")

    def calculate_averages(self, current_record):
        """ 周期汇总统计，加权平均计算年、月度运行指标 """
        target_date = current_record.get('D2', '')
        if not target_date or len(target_date) < 7: return
        t_month, t_year = target_date[:7], target_date[:4]
        m_b, m_c, m_h, y_b, y_c, y_h = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
        for row in self.data_list:
            d2, h = row.get('D2', ''), row.get('H', 0)
            rb, rc = row.get('Rate_Base', 0), row.get('Rate_Comp', 0)
            if h > 0:
                if d2.startswith(t_year): y_b += rb * h; y_c += rc * h; y_h += h
                if d2.startswith(t_month): m_b += rb * h; m_c += rc * h; m_h += h
        _f = lambda s, h: str(round(s/h, 3)) if h > 0 else "0.0"
        self.data_vars['avg_m_base'].set(_f(m_b, m_h)); self.data_vars['avg_m_comp'].set(_f(m_c, m_h))
        self.data_vars['avg_y_base'].set(_f(y_b, y_h)); self.data_vars['avg_y_comp'].set(_f(y_c, y_h))

    def update_display(self):
        """ 将当前索引指向的数据刷新到界面上 """
        if not self.data_list or self.current_idx < 0:
            for v in self.data_vars.values(): v.set("--")
            return
        r, d = self.data_list[self.current_idx], self.data_vars
        mapping = [('P1','p1'),('P2','p2'),('Temp1','temp1'),('Temp2','temp2'),('D1','d1'),('D2','d2'),('T1','t1'),('T2','t2'),('H','h'),('Fill','fill'),('Rate_Base','base'),('Rate_Comp','comp'),('Team','team')]
        for k, vk in mapping:
            d[vk].set(str(r.get(k,'--')))
        self.calculate_averages(r)
        self.page_info.config(text="{0} / {1}".format(self.current_idx + 1, len(self.data_list)))
        for k in ['base', 'comp', 'avg_m_base', 'avg_m_comp', 'avg_y_base', 'avg_y_comp']:
            try:
                val = float(d[k].get())
                self.calc_labels[k].config(fg="#C0392B" if val > 16 else "#27AE60")
            except: self.calc_labels[k].config(fg="#27AE60")

    def switch_unit(self, name):
        """ 切换机组：加载对应的参数配置及历史 JSON 数据 """
        self.current_unit = name; self.title_var.set("机组: {0}".format(name))
        self.params = {"volume": 125.0, "p_atm": 0.101325, "base_temp": 40.0}
        cfg_p = os.path.join(CONFIG_DIR, "{0}.json".format(name))
        if os.path.exists(cfg_p):
            with open(cfg_p, 'r') as f: self.params.update(json.load(f))
        data_p = os.path.join(DATA_DIR, "{0}.json".format(name))
        if os.path.exists(data_p):
            with open(data_p, 'r') as f: self.data_list = json.load(f)
        else: self.data_list = []
        self.recalc_and_save(); self.current_idx = len(self.data_list)-1; self.update_display()

    def recalc_and_save(self):
        """ 核心计算引擎：基于克拉伯龙方程及其变体，实现氢量和漏氢率的自动换算 """
        if not self.data_list: return
        self.data_list.sort(key=lambda x: x.get('D2','') + x.get('T2',''))
        v, p0, bt = self.params['volume'], self.params['p_atm'], self.params['base_temp']
        for i, row in enumerate(self.data_list):
            try:
                # 克拉伯龙方程变体应用
                f = v * ((row['P2']+p0)/(row['Temp2']+273.15) - (row['P1']+p0)/(row['Temp1']+273.15)) * 293.15 / 0.101325
                row['Fill'] = round(f, 2)
                if i > 0:
                    t_s = datetime.strptime(row['D1'] + " " + row['T1'], "%Y-%m-%d %H:%M")
                    t_pe = datetime.strptime(self.data_list[i-1]['D2'] + " " + self.data_list[i-1]['T2'], "%Y-%m-%d %H:%M")
                    dt = (t_s - t_pe).total_seconds() / 3600; row['H'] = round(dt, 2)
                    if dt > 0:
                        lp2, lt2 = self.data_list[i-1]['P2'], self.data_list[i-1]['Temp2']
                        rb = (v*24/dt) * ((lp2+p0)/(lt2+273.15) - (row['P1']+p0)/(row['Temp1']+273.15)) * 293.15 / 0.101325
                        row['Rate_Base'] = round(rb, 3)
                        row['Rate_Comp'] = round(rb * (273.15 + bt) / (273.15 + row['Temp1']), 3)
                else: row['H'], row['Rate_Base'], row['Rate_Comp'] = 0.0, 0.0, 0.0
            except: continue
        with open(os.path.join(DATA_DIR, "{0}.json".format(self.current_unit)), 'w') as f:
            json.dump(self.data_list, f, indent=2)

    def load_units(self):
        """ 加载机组名录 """
        try:
            if os.path.exists(INDEX_FILE):
                with open(INDEX_FILE, 'r') as f: return json.load(f)
            else: return []
        except: return []

    def refresh_unit_sidebar(self):
        """ 重新生成侧边栏的机组选择按钮 """
        for b in self.unit_btns.values(): b.destroy()
        self.unit_btns = {}
        for u in self.units:
            btn = tk.Button(self.unit_scroll_fm, text=u, bg="#34495E", fg="white", font=self.font_normal, command=lambda n=u: self.switch_unit(n))
            btn.pack(fill="x", pady=1, padx=8); self.unit_btns[u] = btn

    def toggle_admin(self):
        """ 权限切换逻辑 """
        if not self.is_admin:
            if simpledialog.askstring("验证", "请输入管理员密码:", show='*') == self.admin_pwd: self.is_admin = True
        else: self.is_admin = False
        self.refresh_ui_state()

    def move_idx(self, step):
        """ 实现记录的翻页查看 """
        if self.data_list: self.current_idx = max(0, min(len(self.data_list)-1, self.current_idx+step)); self.update_display()

    def delete_record(self):
        """ 删除单条历史记录 """
        if self.current_idx >= 0 and messagebox.askyesno("确认", "确定永久删除当前记录?"):
            del self.data_list[self.current_idx]; self.current_idx = max(0, self.current_idx-1); self.recalc_and_save(); self.update_display()

    def open_record_window(self): self._create_record_popup("新增补氢记录")
    def open_edit_window(self):
        if self.current_idx >= 0: self._create_record_popup("修改补氢记录", self.data_list[self.current_idx])

    def open_config_ui(self):
        """ 机组容积等基础物理参数的配置界面 """
        win = tk.Toplevel(self.root); win.title("{0} 参数配置".format(self.current_unit))
        win.geometry("{0}x{1}".format(int(850*self.scale), int(450*self.scale)))
        fm = tk.Frame(win, pady=20, padx=60); fm.pack(fill="both")
        flds = [("机组容积 (m³):", "volume"), ("基准氢温 (℃):", "base_temp"), ("大气压力 (MPa):", "p_atm")]
        ents = {}
        for i, (l, k) in enumerate(flds):
            tk.Label(fm, text=l, font=self.font_normal).grid(row=i, column=0, sticky="w", pady=10)
            e = tk.Entry(fm, font=self.font_normal, width=15); e.insert(0, str(self.params.get(k, ""))); e.grid(row=i, column=1, sticky="w"); ents[k] = e
        def sv():
            try:
                for k in ents: self.params[k] = float(ents[k].get())
                with open(os.path.join(CONFIG_DIR, "{0}.json".format(self.current_unit)), 'w') as f:
                    json.dump(self.params, f)
                self.recalc_and_save(); self.update_display(); win.destroy()
            except: messagebox.showerror("错误", "参数格式错误")
        tk.Button(win, text="保存参数", font=self.font_main, bg="#7F8C8D", fg="white", padx=25, command=sv).pack(pady=10)

    def add_unit_ui(self):
        n = simpledialog.askstring("添加机组", "请输入新机组名称:")
        if n: self.units.append(n); self.save_index(); self.refresh_unit_sidebar()

    def rename_unit_ui(self):
        if self.current_unit:
            n = simpledialog.askstring("重命名", "请输入新的机组名称:")
            if n: 
                idx = self.units.index(self.current_unit); self.units[idx] = n
                self.save_index(); self.refresh_unit_sidebar(); self.switch_unit(n)

    def save_index(self):
        with open(INDEX_FILE, 'w') as f: json.dump(self.units, f)

    def open_export_dialog(self):
        """ 自动化报表生成：基于 xlsxwriter 导出带格式、带计算汇总的 Excel 报表 """
        if not self.data_list: return
        now = datetime.now()
        year = simpledialog.askinteger("导出", "请输入导出年份:", initialvalue=now.year)
        if not year: return
        month = simpledialog.askinteger("导出", "请输入导出月份(1-12):", initialvalue=now.month)
        if not month: return
        
        # 筛选指定月份的数据
        target_month_str = "{0}-{1:02d}".format(year, month)
        filtered_data = [r for r in self.data_list if r['D2'].startswith(target_month_str)]
        if not filtered_data:
            messagebox.showinfo("提示", "未找到该月份的补氢数据记录")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="{0}_{1}漏氢率统计表.xlsx".format(self.current_unit, target_month_str))
        if not file_path: return
        try:
            wb = xlsxwriter.Workbook(file_path)
            ws = wb.add_worksheet("Sheet1")

            # 定义 Excel 单元格格式
            fmt_title = wb.add_format({'bold':True,'font_size':18,'align':'center','valign':'vcenter','border':1})
            fmt_cell = wb.add_format({'align':'center','valign':'vcenter','border':1})
            fmt_red = wb.add_format({'align':'center','valign':'vcenter','border':1,'font_color':'red','bold':True})

            # 定义标题颜色区分格式
            h_base = wb.add_format({'bold':True,'align':'center','valign':'vcenter','bg_color':'#E8E8E8','border':1, 'text_wrap': True}) 
            h_pre = wb.add_format({'bold':True,'align':'center','valign':'vcenter','bg_color':'#FFF2CC','border':1, 'text_wrap': True})  
            h_post = wb.add_format({'bold':True,'align':'center','valign':'vcenter','bg_color':'#DDEBF7','border':1, 'text_wrap': True}) 
            h_res = wb.add_format({'bold':True,'align':'center','valign':'vcenter','bg_color':'#E2EFDA','border':1, 'text_wrap': True})  

            # 设置列宽及大标题
            ws.set_column('A:M', 14); ws.set_row(0, 45) 
            ws.merge_range(0, 0, 0, 12, "{0} {1}年{2}月 漏氢率统计表".format(self.current_unit, year, month), fmt_title)

            # 写入表头
            headers = [("补氢班组", h_base), ("开始日期", h_pre), ("开始时间", h_pre), ("补前压力\n(MPa)", h_pre), ("补前氢温\n(℃)", h_pre), ("结束日期", h_post), ("结束时间", h_post), ("补后压力\n(MPa)", h_post), ("补后氢温\n(℃)", h_post), ("本次补氢量\n(m³)", h_res), ("漏氢时长\n(h)", h_res), ("基准漏氢率\n(m³/d)", h_res), ("补偿漏氢率\n(m³/d)", h_res)]
            ws.set_row(1, 45) 
            for c, header_data in enumerate(headers): ws.write(1, c, header_data[0], header_data[1])

            # 写入历史记录行数据
            curr_row = 2
            for r_data in filtered_data:
                ws.set_row(curr_row, 25)
                ws.write(curr_row, 0, r_data.get('Team',''), fmt_cell)
                ws.write(curr_row, 1, r_data.get('D1',''), fmt_cell); ws.write(curr_row, 2, r_data.get('T1',''), fmt_cell)
                ws.write(curr_row, 3, r_data.get('P1',0), fmt_cell); ws.write(curr_row, 4, r_data.get('Temp1',0), fmt_cell)
                ws.write(curr_row, 5, r_data.get('D2',''), fmt_cell); ws.write(curr_row, 6, r_data.get('T2',''), fmt_cell)
                ws.write(curr_row, 7, r_data.get('P2',0), fmt_cell); ws.write(curr_row, 8, r_data.get('Temp2',0), fmt_cell)
                ws.write(curr_row, 9, r_data.get('Fill',0), fmt_cell); ws.write(curr_row, 10, r_data.get('H',0), fmt_cell)

                # 漏氢率超标时在 Excel 中也应用红色高亮显示
                v_base = r_data.get('Rate_Base', 0); ws.write(curr_row, 11, v_base, fmt_red if v_base > 16 else fmt_cell)
                v_comp = r_data.get('Rate_Comp', 0); ws.write(curr_row, 12, v_comp, fmt_red if v_comp > 16 else fmt_cell)
                curr_row += 1

            # 写入底部周期汇总统计数据
            fmt_sum_label = wb.add_format({'bold':True, 'bg_color':'#F2F2F2', 'border':1, 'align':'center'})
            fmt_sum_val = wb.add_format({'bold':True, 'border':1, 'font_color':'#2E75B6', 'align':'center'})
            curr_row += 2 
            ws.merge_range(curr_row, 0, curr_row, 3, "周期汇总统计 (加权平均值)", fmt_sum_label)
            vals = [("本月基准月均率", self.data_vars['avg_m_base'].get(), "本月补偿月均率", self.data_vars['avg_m_comp'].get()), ("本年基准年均率", self.data_vars['avg_y_base'].get(), "本年补偿年均率", self.data_vars['avg_y_comp'].get())]
            
            for i, val_row in enumerate(vals):
                r = curr_row + 1 + i
                ws.write(r, 0, val_row[0], fmt_sum_label); ws.write(r, 1, val_row[1], fmt_sum_val)
                ws.write(r, 2, val_row[2], fmt_sum_label); ws.write(r, 3, val_row[3], fmt_sum_val)
            wb.close(); messagebox.showinfo("成功", "报表已生成至：\n" + file_path)
        except Exception as e: messagebox.showerror("错误", "导出报表时发生异常: " + str(e))

# 程序执行入口
if __name__ == "__main__":
    root = tk.Tk()
    app = HydrogenAppXP(root)
    root.mainloop()