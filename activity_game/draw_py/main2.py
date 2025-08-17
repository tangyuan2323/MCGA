import pandas as pd
import random
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk, ImageDraw
import os
import sys, os
BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

BG_FILE = os.path.join(BASE_DIR, "8.png")

# ===================== 全局参数 =====================
WIN_W, WIN_H = 1024, 1024          # 窗口 / 背景分辨率
# BG_FILE = "8.png"                  # 背景图片文件名
# ===================================================


class LotteryApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🎉 MCGA抽奖小程序")
        self.root.geometry(f"{WIN_W}x{WIN_H}")
        self.root.resizable(False, False)

        # 数据容器
        self.participants = pd.DataFrame()
        self.drawn_participants = []

        # 背景处理
        self.create_background()
        self.setup_ui()

    # ---------- 背景 ----------
    def create_background(self):
        """优先加载 BG_FILE，不成功则使用渐变背景"""
        try:
            img = Image.open(BG_FILE)
            if img.size != (WIN_W, WIN_H):
                img = img.resize((WIN_W, WIN_H), Image.Resampling.LANCZOS)
            self.bg_image = ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"[INFO] 加载 {BG_FILE} 失败，使用渐变背景 → {e}")
            self.create_gradient_background()

    def create_gradient_background(self):
        """生成彩虹渐变背景"""
        img = Image.new("RGB", (WIN_W, WIN_H))
        draw = ImageDraw.Draw(img)

        for y in range(WIN_H):
            t = y / WIN_H
            if t < 0.25:      # 深紫 → 紫
                r = int(75 + (138 - 75) * t / 0.25)
                g = int(0 + (43 - 0) * t / 0.25)
                b = int(130 + (226 - 130) * t / 0.25)
            elif t < 0.5:     # 紫 → 蓝
                r = int(138 + (64 - 138) * (t - .25) / .25)
                g = int(43 + (224 - 43) * (t - .25) / .25)
                b = int(226 + (208 - 226) * (t - .25) / .25)
            elif t < 0.75:    # 蓝 → 青
                r = int(64 + (72 - 64) * (t - .5) / .25)
                g = int(224 + (209 - 224) * (t - .5) / .25)
                b = int(208 + (204 - 208) * (t - .5) / .25)
            else:             # 青 → 绿
                r = int(72 + (34 - 72) * (t - .75) / .25)
                g = int(209 + (197 - 209) * (t - .75) / .25)
                b = int(204 + (94 - 204) * (t - .75) / .25)
            draw.line([(0, y), (WIN_W, y)], fill=(r, g, b))

        # 星星点缀
        for _ in range(20):
            x, y = random.randint(20, WIN_W - 20), random.randint(20, WIN_H - 20)
            self.draw_star(draw, x, y, 8, (255, 255, 255))

        self.bg_image = ImageTk.PhotoImage(img)

    def draw_star(self, draw: ImageDraw.ImageDraw, cx, cy, size, color):
        """在给定 Draw 对象上画五角星"""
        import math
        pts = []
        for i in range(10):
            ang = i * math.pi / 5
            r = size if i % 2 == 0 else size / 2
            pts.append((cx + r * math.cos(ang - math.pi / 2),
                        cy + r * math.sin(ang - math.pi / 2)))
        draw.polygon(pts, fill=color)

    # ---------- UI ----------
    def setup_ui(self):
        # 背景
        tk.Label(self.root, image=self.bg_image).place(x=0, y=0, relwidth=1, relheight=1)

        main = tk.Frame(self.root, bg="#f8f9fa", bd=2, relief="ridge")
        main.place(relx=0.1, rely=0.05, relwidth=0.8, relheight=0.9)

        tk.Label(main, text="🎉 MCGA幸运抽奖大转盘 🎉",
                 font=("Microsoft YaHei", 28, "bold"),
                 fg="#2c3e50", bg="#f8f9fa").pack(pady=25)

        tk.Label(main, text="✨ 公平 · 公正 · 公开 ✨",
                 font=("Microsoft YaHei", 13),
                 fg="#7f8c8d", bg="#f8f9fa").pack()

        # 按钮区
        btn_frame = tk.Frame(main, bg="#f8f9fa")
        btn_frame.pack(pady=25)

        tk.Button(btn_frame, text="📁 选择Excel文件", width=18, height=2,
                  font=("Microsoft YaHei", 12, "bold"),
                  bg="#3498db", fg="white",
                  activebackground="#2980b9",
                  command=self.load_excel).pack(side="left", padx=15)

        self.lottery_btn = tk.Button(btn_frame, text="🎲 开始抽奖", width=18, height=2,
                                     font=("Microsoft YaHei", 12, "bold"),
                                     bg="#e74c3c", fg="white",
                                     activebackground="#c0392b",
                                     state="disabled",
                                     command=self.start_lottery)
        self.lottery_btn.pack(side="left", padx=15)

        # 提示信息
        self.info_lbl = tk.Label(main,
                                 text="🎯 请先选择包含姓名和编号的 Excel 文件",
                                 font=("Microsoft YaHei", 12),
                                 bg="#f8f9fa", fg="#34495e")
        self.info_lbl.pack(pady=10)

        # 结果框
        res_labelf = tk.LabelFrame(main, text="🏆 中奖结果",
                                   font=("Microsoft YaHei", 14, "bold"),
                                   bg="#f8f9fa", fg="#2c3e50", bd=2, relief="groove")
        res_labelf.pack(fill="both", expand=True, padx=25, pady=15)

        text_wrap = tk.Frame(res_labelf, bg="#f8f9fa")
        text_wrap.pack(fill="both", expand=True, padx=10, pady=10)

        self.result_txt = tk.Text(text_wrap, font=("Consolas", 12),
                                  bg="white", fg="#2c3e50", wrap="word")
        self.result_txt.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(text_wrap, command=self.result_txt.yview)
        sb.pack(side="right", fill="y")
        self.result_txt.config(yscrollcommand=sb.set)

        # 控制按钮
        ctrl = tk.Frame(main, bg="#f8f9fa")
        ctrl.pack(pady=10)

        self.reset_btn = tk.Button(ctrl, text="🔄 重置抽奖", width=12,
                                   font=("Microsoft YaHei", 10, "bold"),
                                   bg="#f39c12", fg="white",
                                   activebackground="#e67e22",
                                   state="disabled",
                                   command=self.reset_lottery)
        self.reset_btn.pack(side="left", padx=15)

        self.save_btn = tk.Button(ctrl, text="💾 保存结果", width=12,
                                  font=("Microsoft YaHei", 10, "bold"),
                                  bg="#27ae60", fg="white",
                                  activebackground="#229954",
                                  state="disabled",
                                  command=self.save_results)
        self.save_btn.pack(side="left", padx=15)

        # 初始内容
        self.result_txt.insert("1.0",
            "🎪 欢迎使用MCGA抽奖系统！\n\n"
            "📋 步骤：\n"
            "1. 点击“选择Excel文件”加载名单\n"
            "2. 点击“开始抽奖”\n"
            "3. 可以多轮抽奖，已中奖者不重复\n\n"
            "Excel 必须包含：\n"
            "• 姓名列（姓名/name/Name/名字）\n"
            "• 编号列（编号/id/ID/号码/number）\n")

    # ---------- 功能 ----------
    def load_excel(self):
        path = filedialog.askopenfilename(title="选择 Excel 文件",
                                          filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return

        name_col = next((c for c in df.columns
                         if c in ('姓名', 'name', 'Name', '名字')), None)
        id_col   = next((c for c in df.columns
                         if c in ('编号', 'id', 'ID', 'Id', '号码', 'number')), None)
        if not name_col or not id_col:
            messagebox.showerror("列缺失", "Excel 必须同时包含姓名列和编号列！")
            return

        self.participants = df[[name_col, id_col]].dropna().rename(
            columns={name_col: '姓名', id_col: '编号'})
        self.drawn_participants.clear()

        self.info_lbl.config(text=f"✅ 已加载 {len(self.participants)} 位参与者")
        self.lottery_btn.config(state="normal")
        self.reset_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.result_txt.delete("1.0", "end")
        self.result_txt.insert("1.0", "准备就绪，点击“开始抽奖”！")

    def start_lottery(self):
        if self.participants.empty:
            messagebox.showwarning("未加载", "请先加载 Excel 文件")
            return

        remaining_df = self.participants[~self.participants['编号'].isin(
            [p['编号'] for p in self.drawn_participants])]
        remain = len(remaining_df)
        if remain == 0:
            messagebox.showinfo("结束", "所有人都已中奖！")
            return

        cnt = simpledialog.askinteger("抽奖人数",
                                      f"剩余 {remain} 人，输入要抽取的人数：",
                                      minvalue=1, maxvalue=remain)
        if not cnt:
            return

        winners = remaining_df.sample(cnt)
        for _, row in winners.iterrows():
            self.drawn_participants.append(row.to_dict())

        self.display_results()
        self.reset_btn.config(state="normal")
        self.save_btn.config(state="normal")

    def display_results(self):
        self.result_txt.delete("1.0", "end")
        self.result_txt.insert("end", f"🎉 当前已产生 {len(self.drawn_participants)} 位中奖者：\n\n")
        for i, p in enumerate(self.drawn_participants, 1):
            self.result_txt.insert("end", f"{i:02d}. {p['姓名']}  (编号: {p['编号']})\n")

        remain = len(self.participants) - len(self.drawn_participants)
        self.result_txt.insert("end", f"\n剩余未中奖人数：{remain}\n")

    def reset_lottery(self):
        self.drawn_participants.clear()
        self.result_txt.delete("1.0", "end")
        self.result_txt.insert("1.0", "🎲 已重置，点击“开始抽奖”重新开始！")
        self.reset_btn.config(state="disabled")
        self.save_btn.config(state="disabled")

    def save_results(self):
        if not self.drawn_participants:
            messagebox.showwarning("无结果", "没有可保存的中奖结果")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel 文件", "*.xlsx")],
                                            title="保存抽奖结果")
        if not path:
            return
        try:
            pd.DataFrame(self.drawn_participants).to_excel(path, index=False)
            messagebox.showinfo("保存成功", f"已保存到：{path}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))


def main():
    root = tk.Tk()
    LotteryApp(root)

    # 若有图标文件
    if os.path.exists("lottery.ico"):
        try:
            root.iconbitmap("lottery.ico")
        except:
            pass

    root.eval('tk::PlaceWindow . center')
    root.mainloop()


if __name__ == "__main__":
    main()