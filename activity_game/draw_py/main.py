import pandas as pd
import random
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os

class LotteryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("抽奖小程序")
        self.root.geometry("500x400")
        
        # 数据存储
        self.participants = pd.DataFrame()
        self.drawn_participants = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # 标题
        title_label = tk.Label(self.root, text="🎉 抽奖小程序 🎉", 
                              font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 文件选择按钮
        file_button = tk.Button(self.root, text="选择Excel文件", 
                               command=self.load_excel, 
                               font=("Arial", 12),
                               bg="#4CAF50", fg="white",
                               width=20, height=2)
        file_button.pack(pady=10)
        
        # 显示参与人数
        self.info_label = tk.Label(self.root, text="请先选择Excel文件", 
                                  font=("Arial", 10))
        self.info_label.pack(pady=5)
        
        # 抽奖按钮
        self.lottery_button = tk.Button(self.root, text="开始抽奖", 
                                       command=self.start_lottery,
                                       font=("Arial", 14, "bold"),
                                       bg="#FF5722", fg="white",
                                       width=20, height=2,
                                       state="disabled")
        self.lottery_button.pack(pady=20)
        
        # 结果显示区域
        result_frame = tk.Frame(self.root)
        result_frame.pack(pady=10)
        
        tk.Label(result_frame, text="中奖结果：", font=("Arial", 12, "bold")).pack()
        
        self.result_text = tk.Text(result_frame, width=50, height=8, 
                                  font=("Arial", 11))
        self.result_text.pack()
        
        # 控制按钮
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.reset_button = tk.Button(button_frame, text="重置抽奖", 
                                     command=self.reset_lottery,
                                     font=("Arial", 10),
                                     bg="#FFC107",
                                     state="disabled")
        self.reset_button.pack(side="left", padx=5)
        
        self.save_button = tk.Button(button_frame, text="保存结果", 
                                    command=self.save_results,
                                    font=("Arial", 10),
                                    bg="#2196F3", fg="white",
                                    state="disabled")
        self.save_button.pack(side="left", padx=5)
    
    def load_excel(self):
        """加载Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            try:
                # 读取Excel文件
                self.participants = pd.read_excel(file_path)
                
                # 检查列名（支持多种可能的列名）
                name_cols = ['姓名', 'name', 'Name', '名字']
                id_cols = ['编号', 'id', 'ID', 'Id', '号码', 'number']
                
                name_col = None
                id_col = None
                
                for col in self.participants.columns:
                    if col in name_cols:
                        name_col = col
                    if col in id_cols:
                        id_col = col
                
                if name_col is None or id_col is None:
                    messagebox.showerror("错误", 
                                       "Excel文件必须包含姓名和编号列！\n"
                                       "支持的列名：\n"
                                       "姓名：姓名/name/Name/名字\n"
                                       "编号：编号/id/ID/Id/号码/number")
                    return
                
                # 标准化列名
                self.participants = self.participants[[name_col, id_col]]
                self.participants.columns = ['姓名', '编号']
                
                # 去除空值
                self.participants = self.participants.dropna()
                
                participant_count = len(self.participants)
                self.info_label.config(text=f"已加载 {participant_count} 位参与者")
                
                # 启用抽奖按钮
                self.lottery_button.config(state="normal")
                self.reset_lottery()
                
                messagebox.showinfo("成功", f"成功加载 {participant_count} 位参与者！")
                
            except Exception as e:
                messagebox.showerror("错误", f"读取文件失败：{str(e)}")
    
    def start_lottery(self):
        """开始抽奖"""
        if self.participants.empty:
            messagebox.showwarning("警告", "请先加载Excel文件！")
            return
        
        # 获取抽奖人数
        available_count = len(self.participants) - len(self.drawn_participants)
        if available_count <= 0:
            messagebox.showinfo("提示", "所有人都已被抽中！")
            return
        
        draw_count = simpledialog.askinteger(
            "抽奖人数", 
            f"请输入要抽取的人数（剩余{available_count}人）：",
            minvalue=1, 
            maxvalue=available_count
        )
        
        if draw_count:
            # 获取未被抽中的参与者
            remaining_participants = self.participants[
                ~self.participants['编号'].isin([p['编号'] for p in self.drawn_participants])
            ]
            
            # 随机抽取
            winners = remaining_participants.sample(n=draw_count)
            
            # 添加到已抽奖列表
            for _, winner in winners.iterrows():
                self.drawn_participants.append({
                    '姓名': winner['姓名'],
                    '编号': winner['编号']
                })
            
            # 显示结果
            self.display_results()
            
            # 启用相关按钮
            self.reset_button.config(state="normal")
            self.save_button.config(state="normal")
    
    def display_results(self):
        """显示抽奖结果"""
        self.result_text.delete(1.0, tk.END)
        
        if not self.drawn_participants:
            self.result_text.insert(tk.END, "暂无抽奖结果")
            return
        
        result_text = f"🎊 恭喜以下 {len(self.drawn_participants)} 位中奖者 🎊\n"
        result_text += "=" * 40 + "\n"
        
        for i, participant in enumerate(self.drawn_participants, 1):
            result_text += f"{i:2d}. {participant['姓名']} (编号: {participant['编号']})\n"
        
        result_text += "=" * 40 + "\n"
        result_text += f"剩余参与者：{len(self.participants) - len(self.drawn_participants)} 人"
        
        self.result_text.insert(tk.END, result_text)
    
    def reset_lottery(self):
        """重置抽奖"""
        self.drawn_participants = []
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "点击'开始抽奖'进行抽奖")
        self.reset_button.config(state="disabled")
        self.save_button.config(state="disabled")
    
    def save_results(self):
        """保存抽奖结果"""
        if not self.drawn_participants:
            messagebox.showwarning("警告", "暂无抽奖结果可保存！")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存抽奖结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                results_df = pd.DataFrame(self.drawn_participants)
                results_df.to_excel(file_path, index=False)
                messagebox.showinfo("成功", f"抽奖结果已保存到：{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败：{str(e)}")

def main():
    root = tk.Tk()
    app = LotteryApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()