import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk
import pandas as pd
from typing import List, Dict, Any

class MaterialityAssessment:
    ITEMS = [
        "永續策略", "誠信經營", "公司治理", "稅務政策", "風險控管", "法規遵循",
        "營運持續管理", "營運績效", "創新與數位責任", "資訊安全", "供應商管理",
        "客戶關係管理", "氣候變遷因應", "能資源管理", "職場健康與安全", "員工培育與職涯發展",
        "人才吸引與留任", "社會關懷與鄰里促進", "人權平等"
    ]
    
    ISSUE_TYPES = ["實際", "潛在"]
    FORM_FIELDS = ["議題類型", "機會", "機會實現可能性", "風險議題", "風險發生可能性"]
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("永續發展重大性評估")
        self.root.geometry("800x700")
        
        self.setup_font()
        self.selected_items: List[str] = []
        self.item_vars: List[tk.BooleanVar] = []
        
        # Create name and department fields at the top
        self.create_header_frame()
        self.setup_scrollable_canvas()
        self.create_selection_page()
        
    def setup_font(self):
        if self.root.tk.call("tk", "windowingsystem") == "win32":
            self.chinese_font = font.Font(family="Microsoft JhengHei", size=10)
        elif self.root.tk.call("tk", "windowingsystem") == "aqua":
            self.chinese_font = font.Font(family="PingFang TC", size=10)
        else:
            self.chinese_font = font.Font(family="Noto Sans CJK", size=10)
    
    def create_header_frame(self):
        self.header_frame = tk.Frame(self.root)
        self.header_frame.pack(fill="x", padx=10, pady=5)
        
        # Name field
        name_label = tk.Label(self.header_frame, text="姓名：", font=self.chinese_font)
        name_label.pack(side="left", padx=(0, 5))
        self.name_var = tk.StringVar()
        name_entry = tk.Entry(self.header_frame, textvariable=self.name_var, font=self.chinese_font, width=15)
        name_entry.pack(side="left", padx=(0, 20))
        
        # Department field
        dept_label = tk.Label(self.header_frame, text="部門：", font=self.chinese_font)
        dept_label.pack(side="left", padx=(0, 5))
        self.department_var = tk.StringVar()
        dept_entry = tk.Entry(self.header_frame, textvariable=self.department_var, font=self.chinese_font, width=15)
        dept_entry.pack(side="left")
    
    def setup_scrollable_canvas(self):
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.frame = tk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def on_checkbutton_select(self):
        selected = [i for i, var in enumerate(self.item_vars) if var.get()]
        if len(selected) > 10:
            messagebox.showwarning("警告", "最多只能選擇10個項目！")
            self.item_vars[selected[-1]].set(False)
            return
            
        self.selected_items.clear()
        self.selected_items.extend([self.ITEMS[i] for i in selected])
    
    def create_selection_page(self):
        # Clear existing widgets in frame
        for widget in self.frame.winfo_children():
            widget.destroy()
        
        # Clear item vars and recreate
        self.item_vars.clear()
        
        tk.Label(self.frame, text="請選擇10個重大性評估項目", font=self.chinese_font).grid(
            row=0, column=0, columnspan=2, pady=10)
        
        # Create checkbuttons
        for i, item in enumerate(self.ITEMS):
            var = tk.BooleanVar()
            # If item was previously selected, check it
            if item in self.selected_items:
                var.set(True)
            cb = tk.Checkbutton(
                self.frame, 
                text=item, 
                variable=var, 
                font=self.chinese_font,
                command=self.on_checkbutton_select
            )
            cb.grid(row=i+1, column=0, columnspan=2, sticky="w", padx=10, pady=5)
            self.item_vars.append(var)
        
        # Submit button
        submit_frame = tk.Frame(self.root)
        submit_frame.pack(fill="x", pady=10)
        tk.Button(
            submit_frame,
            text="提交選擇並生成問卷",
            font=self.chinese_font,
            command=self.generate_questionnaire
        ).pack(pady=5)
    
    def create_button_frame(self):
        if hasattr(self, 'button_frame'):
            self.button_frame.destroy()
        
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(fill="x", pady=10)
        
        # Back button
        tk.Button(
            self.button_frame,
            text="返回修改選項",
            font=self.chinese_font,
            command=self.create_selection_page
        ).pack(side="left", padx=10)
        
        # Save button (if on questionnaire page)
        if hasattr(self, 'form_data'):
            tk.Button(
                self.button_frame,
                text="保存結果",
                font=self.chinese_font,
                command=self.save_results
            ).pack(side="right", padx=10)
    
    def generate_questionnaire(self):
        if not self.name_var.get() or not self.department_var.get():
            messagebox.showwarning("警告", "請填寫姓名和部門！")
            return
            
        if len(self.selected_items) != 10:
            messagebox.showwarning("警告", "必須選擇10個項目！")
            return
            
        # Clear existing widgets
        for widget in self.frame.winfo_children():
            widget.destroy()
            
        tk.Label(self.frame, text="風險與機會評估表", font=self.chinese_font).grid(
            row=0, column=0, columnspan=2, pady=10)
        
        self.form_data = []
        row = 1
        
        for i, item in enumerate(self.selected_items):
            tk.Label(
                self.frame,
                text=f"項目 {i+1}：{item}",
                font=self.chinese_font,
                bg="lightgray",
                width=40
            ).grid(row=row, column=0, columnspan=2, padx=10, pady=5, sticky="w")
            row += 1
            
            item_data = []
            for j, label in enumerate(self.FORM_FIELDS):
                tk.Label(self.frame, text=label, font=self.chinese_font).grid(
                    row=row, column=0, padx=10, pady=5, sticky="e")
                
                if j == 0:  # For the first field (議題類型)
                    var = tk.StringVar(value=self.ISSUE_TYPES[0])
                    dropdown = ttk.Combobox(
                        self.frame,
                        textvariable=var,
                        values=self.ISSUE_TYPES,
                        state="readonly",
                        width=10
                    )
                    dropdown.grid(row=row, column=1, padx=10, pady=5, sticky="w")
                    item_data.append(var)
                else:  # For other fields (scales)
                    entry_var = tk.IntVar(value=3)
                    scale = tk.Scale(
                        self.frame,
                        from_=1,
                        to=5,
                        orient="horizontal",
                        variable=entry_var,
                        length=200
                    )
                    scale.grid(row=row, column=1, padx=10, pady=5, sticky="w")
                    item_data.append(entry_var)
                
                row += 1
                
            self.form_data.append(item_data)
        
        # Update button frame with both back and save buttons
        self.create_button_frame()
        
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def save_results(self):
        data = []
        for idx, item in enumerate(self.selected_items):
            item_data = self.form_data[idx]
            data.append({
                "項目": item,
                "議題類型": item_data[0].get(),
                "機會": item_data[1].get(),
                "機會實現可能性": item_data[2].get(),
                "風險議題": item_data[3].get(),
                "風險發生可能性": item_data[4].get(),
            })
            
        df = pd.DataFrame(data)
        save_path = filedialog.askdirectory()
        if not save_path:
            return
            
        file_name = f"materiality-assessed-{self.name_var.get()}.csv"
        full_path = f"{save_path}/{file_name}"
        df.to_csv(full_path, index=False, encoding='utf-8-sig')
        
        messagebox.showinfo("保存成功", f"資料已保存至: {full_path}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MaterialityAssessment()
    app.run()