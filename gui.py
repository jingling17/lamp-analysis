import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from lamp_analysis import LampAnalysis

class LampAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("台灯销售数据分析工具")
        self.root.geometry("600x400")
        
        # 创建主框架
        self.main_frame = tk.Frame(self.root, padx=20, pady=20)
        self.main_frame.pack(expand=True, fill='both')
        
        # 标题
        title_label = tk.Label(self.main_frame, text="台灯销售数据分析工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        # 文件选择区域
        self.file_frame = tk.Frame(self.main_frame)
        self.file_frame.pack(fill='x', pady=20)
        
        self.file_path_var = tk.StringVar()
        self.file_path_entry = tk.Entry(self.file_frame, textvariable=self.file_path_var, width=50)
        self.file_path_entry.pack(side='left', padx=5)
        
        self.select_button = tk.Button(self.file_frame, text="选择文件", command=self.select_file)
        self.select_button.pack(side='left', padx=5)
        
        # 分析按钮
        self.analyze_button = tk.Button(self.main_frame, text="开始分析", command=self.run_analysis,
                                      width=20, height=2)
        self.analyze_button.pack(pady=20)
        
        # 状态显示
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(self.main_frame, textvariable=self.status_var,
                                   wraplength=500, justify='left')
        self.status_label.pack(pady=20)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            
    def run_analysis(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("错误", "请先选择Excel文件！")
            return
            
        try:
            self.status_var.set("正在分析数据...")
            self.root.update()
            
            # 运行分析
            analyzer = LampAnalysis(file_path)
            analyzer.add_price_range()
            analyzer.analyze_total_sales()
            analyzer.analyze_price_range_distribution()
            analyzer.analyze_brand_market_share()
            analyzer.analyze_top_brands_price_distribution()
            analyzer.save_analysis_to_excel()
            
            self.status_var.set("分析完成！\n报告已保存为：台灯销售分析报告.xlsx\n图表已保存在当前目录下。")
            messagebox.showinfo("完成", "数据分析已完成！")
            
        except Exception as e:
            self.status_var.set(f"分析过程中出现错误：{str(e)}")
            messagebox.showerror("错误", f"分析失败：{str(e)}")

def main():
    root = tk.Tk()
    app = LampAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()