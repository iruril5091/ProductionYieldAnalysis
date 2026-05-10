import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Listbox
from ftplib import FTP
import os
from tkinter import ttk
import re
from tabulate import tabulate

class generate_defect_summary_report:
    def __init__(self, master):
        self.master = master
        self.master.title("ProductionYieldAnalysis_251014v1")
        self.master.geometry("900x900")

    # 創建Notebook 分頁
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(expand=True, fill='both')


    # 分頁1：Download檔案管理界面
        self.page_download = ttk.Frame(self.notebook)
        self.notebook.add(self.page_download, text='Download FTP Files')
        self.FTP_RTP_Path = "/home/nfs/ledfs/UMRTP100/DATA/UPLOAD/HIS"
        self.FTP_AOI600_Path = "/home/nfs/ledfs/UMAOI600/DATA/UPLOAD/REAL"
        self.FTP_CT2_Path = "/home/nfs/ledfs/UACT2100/DATA/UPLOAD/HIS"
        self.FTP_AOI100_Path = "/home/nfs/ledfs/UAAOI100/DATA/UPLOAD/HIS"

        self.create_download_page()
        
    # 分頁2：RTP資料瀏覽界面
        self.page_display = ttk.Frame(self.notebook)
        self.notebook.add(self.page_display, text='RTP')
        self.dataframe = None
        self.matching_files = []  # 用於存儲匹配的檔案名
        self.create_RTP_page()

    # 分頁3：MT Yield數據
        self.page_MT_Yield = ttk.Frame(self.notebook)
        self.notebook.add(self.page_MT_Yield, text="MT Yield")
        self.directory = None
        self.csv_file = None
        self.create_MT_Yield_page()

    # 分頁4：ASM Yield數據
        self.page_ASM_Yield = ttk.Frame(self.notebook)
        self.notebook.add(self.page_ASM_Yield, text="ASM Yield")
        self.CT2_file_path = None
        self.create_ASM_Yield_page()

    # 分頁5：PN Yield數據
        self.page_PN_Yield = ttk.Frame(self.notebook)
        self.notebook.add(self.page_PN_Yield, text="PN Yield")
        self.PN_csv_file = None
        self.merged_df = None
        self.create_PN_Yield_page()
#================================================================ 分頁1 Download Page ================================================================#
    def create_download_page(self):
        # 1. 佈局設定用一個框架來排版
        top_frame = tk.Frame(self.page_download)
        top_frame.pack(pady=10, fill='x')

        # 2. FTP路徑設定
        self.ftp_path_var = tk.StringVar()
        self.ftp_path_combobox = ttk.Combobox(top_frame, width=60, textvariable=self.ftp_path_var)
        self.ftp_path_combobox['values'] = [self.FTP_RTP_Path, self.FTP_AOI600_Path, self.FTP_CT2_Path, self.FTP_AOI100_Path]
        self.ftp_path_combobox.pack(side=tk.LEFT, padx=5)

        self.path_button = tk.Button(top_frame, text="設置 FTP 路徑", command=self.update_ftp_path)
        self.path_button.pack(side=tk.LEFT, padx=5)

        # 3. 搜索框與按鈕
        search_frame = tk.Frame(self.page_download)
        search_frame.pack(pady=10, fill='x')

        self.search_entry = tk.Entry(search_frame, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_button = tk.Button(search_frame, text='搜尋FTP檔案', command=self.search_ftp_files)
        self.search_button.pack(side=tk.LEFT, padx=5)

        # 4. 文件清單框（放大一點）
        self.file_listbox = Listbox(self.page_download, width=60, height=10, selectmode='extended')
        self.file_listbox.pack(padx=10, pady=10, fill='both', expand=True)

        # 5. 下載按鈕（放在底部，禁用起來）
        self.download_btn = tk.Button(self.page_download, text='下載選擇檔案', command=self.download_files, state=tk.DISABLED)
        self.download_btn.pack(pady=10)
#================================================================ 分頁1 Download Page Function ================================================================#
    def update_ftp_path(self):
        self.FTP_Path = self.ftp_path_var.get()
        messagebox.showinfo("成功", f"FTP 路徑已更新為: {self.FTP_Path}")

    def search_ftp_files(self):
        search_string = self.search_entry.get()
        try:
            ftp = FTP("10.88.112.25")
            ftp.login("uledeqpftp", "uledeqpftp")
            ftp.cwd(self.FTP_Path)
            filenames = ftp.nlst()

            # 篩選符合搜索字串的檔案
            self.matching_files = [f for f in filenames if search_string in f]

            def extract_date_from_filename(filename):
                match = re.search(r'_(\d{12})', filename)
                if match:
                    # print(f"匹配到的日期：{match.group(1)}")  # 排查用
                    return match.group(1)
                else:
                    # print(f"未找到日期：{filename}")
                    return ""

            # 按照日期由新到舊排序
            self.matching_files.sort(key=extract_date_from_filename, reverse=True)

            # 清空列表框
            self.file_listbox.delete(0, tk.END)

            # 顯示排序後的結果
            if self.matching_files:
                for filename in self.matching_files:
                    self.file_listbox.insert(tk.END, filename)
                self.download_btn.config(state=tk.NORMAL)
                messagebox.showinfo("成功", f"找到以下包含 '{search_string}' 的檔案。")
            else:
                messagebox.showinfo("結果", f"在目錄中沒有找到包含 '{search_string}' 的檔案。")

            ftp.quit()

        except Exception as e:
            messagebox.showerror("錯誤", f"連接到 FTP 伺服器時發生錯誤: {e}")

    def download_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "請選擇檔案！")
            return

        # 選擇存放資料夾
        save_dir = filedialog.askdirectory()
        if not save_dir:
            return

        try:
            ftp = FTP("10.88.112.25")
            ftp.login("uledeqpftp", "uledeqpftp")
            ftp.cwd(self.FTP_Path)

            # 逐個檔案下載
            for index in selected_indices:
                filename = self.matching_files[index]
                local_path = os.path.join(save_dir, filename)

                with open(local_path, 'wb') as file:
                    ftp.retrbinary(f'RETR {filename}', file.write)

            messagebox.showinfo("成功", f"已下載 {len(selected_indices)} 個檔案到資料夾")
            ftp.quit()
        except Exception as e:
            messagebox.showerror("錯誤", f"檔案下載時發生錯誤：{e}")
#================================================================ 分頁2 RTP Page ================================================================#
    def create_RTP_page(self):
        # 1. 建立一個主框架，讓裝置元素更有條理
        frame = tk.Frame(self.page_display)
        frame.pack(padx=10, pady=10, fill='both', expand=True)

        # 2. 文字區域（佔最大片段）
        self.text_area = scrolledtext.ScrolledText(frame, width=90, height=20)
        self.text_area.pack(side=tk.TOP, fill='both', expand=True)

        # 3. 按鈕區域（在文本區下方）
        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=5)

        self.load_button = tk.Button(btn_frame, text='載入 CSV', command=self.load_file)
        self.load_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(btn_frame, text='保存數據', command=self.save_file)
        self.save_button.pack(side=tk.LEFT, padx=5)
        self.save_button.config(state=tk.DISABLED)
#================================================================ 分頁2 RTP Page Function ================================================================#
    def display_data(self, data):
        # 文本框
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, data.to_string(index=False, header=False))
        self.text_area.config(state=tk.DISABLED)

    def load_file(self):
        # 選擇檔案並讀取數據
        file_path = filedialog.askopenfilename(title="載入 CSV 檔案", filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
        if file_path:
            try:    
                df = pd.read_csv(file_path, skiprows=9, header=None)
                clean_df = df.dropna(axis=1, inplace=False)
                clean_df.columns = ["date", "MES_ID", "TOOL_ID", "MODEL_NO", "X", "Y", "gap", "unknow1", "unknow2", "unknow3"]
                clean_df["gap"] = clean_df["gap"].astype(float)

                chip_ID = ["74", "73", "72", "71", "14", "13", "12", "11"]
                clean_df.loc[:, "chip_ID"] = np.repeat(chip_ID, 5)[:len(clean_df)]

                # 合規數據
                range_values = clean_df[(clean_df["gap"] >= 18000) & (clean_df["gap"] < 30000)]
                
                
                range_values.loc[:, 'gap_mean'] = range_values.groupby('chip_ID')['gap'].transform(lambda x: round(x.mean() / 1000, 2))

                range_values['gap_mean'] = range_values.apply(
                    lambda x: x['gap_mean'] if x.name == range_values[range_values['chip_ID'] == x['chip_ID']].index[-1] else None, axis=1)
                
                lock_chip = ["11", "14", "71", "74"]
                print(range_values.loc[(range_values["chip_ID"].isin(lock_chip))])
                
                self.dataframe = range_values

                # 在文本框中顯示 range_values
                self.display_data(range_values.loc[(range_values["chip_ID"].isin(lock_chip))])
                self.save_button.config(state=tk.NORMAL)  # 啟用保存按鈕
                messagebox.showinfo("成功", "檔案載入成功!")
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取檔案時發生錯誤: {e}")

    def save_file(self):
        if self.dataframe is None:
            messagebox.showwarning("警告", "沒有可保存的數據！")
            return      

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
        if save_path:
            try:
                self.dataframe.to_csv(save_path, header=False, index=False, encoding="utf-8")
                messagebox.showinfo("成功", "檔案已保存！")
            except Exception as e:
                messagebox.showerror("錯誤", f"保存檔案時發生錯誤: {e}")
#================================================================ 分頁3 MT_Yield_page ================================================================#
    def create_MT_Yield_page(self):
        thr_frame = tk.Frame(self.page_MT_Yield)
        thr_frame.pack(padx=10, pady=10, fill='both', expand=True)

        # 路徑Entry
        self.MT_dir_entry = tk.Entry(thr_frame, width=60)
        self.MT_dir_entry.pack()
        self.dir_button = tk.Button(thr_frame, text="選擇資料夾路徑", command=self.select_directory)
        self.dir_button.pack()
        
        self.MT_file_list = tk.Listbox(thr_frame,width=80, height=15)
        self.MT_file_list.pack(padx=10, pady=10)

        # text文本
        self.analysis_MT_text = scrolledtext.ScrolledText(thr_frame, width=80, height=20)
        self.analysis_MT_text.pack(padx=10, pady=10)

        # 執行分析按鈕
        run_button = tk.Button(thr_frame, text="執行", command=self.run_MT_analysis)
        run_button.pack(pady=5)
#================================================================ 分頁3 MT_Yield_page Function================================================================#
    def select_directory(self):
        self.directory = filedialog.askdirectory()
        self.csv_file = os.listdir(self.directory)
        if self.directory:
            self.MT_dir_entry.config(state="normal") # 更改為可寫入
            self.MT_dir_entry.delete(0, tk.END) # 清空 Entry
            self.MT_dir_entry.insert(0, self.directory) # 在 Entry 中顯示選擇的路徑
            self.MT_dir_entry.config(text=self.directory) # 顯示
            self.MT_dir_entry.config(state="disabled") # 更改為不可寫入
            self.MT_file_list.delete(0, tk.END)
            for csv in self.csv_file:
                self.MT_file_list.insert(tk.END, csv)

    def run_MT_analysis(self):
        dir_path = self.directory
        if not dir_path:
            messagebox.showwarning("警告", "請先選擇資料夾路徑")
            return
        
        # 1. 取得所有csv檔
        files = [os.path.join(dir_path, f) for f in self.csv_file if f.endswith('_sum.csv')]


        df_list = []
        for file in files:
            df = pd.read_csv(file)
            df_clean = df.dropna(subset=["ShiftX"])
            df_list.append(df_clean)
        if not df_list:
            self.analysis_MT_text.delete(1.0, tk.END)
            self.analysis_MT_text.insert(tk.END, "未找到資料或資料為空。\n")
            return
        df_concat = pd.concat(df_list, ignore_index=True)

        # 2. 進行分析
        total_watch = len(df_concat[df_concat["LED Type"] == "A"])
        O_watch = len(df_concat[(df_concat["LED Type"] == "A") & (df_concat["Judge"] == "O")])
        X_watch = len(df_concat[(df_concat["LED Type"] == "A") & (df_concat["Judge"] == "X")])
        MD_yield = (O_watch / total_watch) * 100 if total_watch > 0 else 0
        sheet_ids = df_concat["#Sheet ID"].unique().tolist()

        # 3. 顯示結果
        result_str = (
            f"Total Watch數量: {total_watch}\n"
            f"O品 Watch數量: {O_watch}\n"
            f"X品 Watch數量: {X_watch}\n"
            f"Yield: {MD_yield:.1f}%\n\n"
            f"Sheet ID:\n" + "\n".join(sheet_ids)
        )
        self.analysis_MT_text.delete(1.0, tk.END)
        self.analysis_MT_text.insert(tk.END, result_str)
#================================================================ 分頁4 ASM_Yield_page ================================================================#
    def create_ASM_Yield_page(self):
        frame4 = tk.Frame(self.page_ASM_Yield)
        frame4.pack(padx=10, pady=10, fill='both', expand=True)

        # 上方輸入與按鈕區域
        top_frame = tk.Frame(frame4)
        top_frame.pack(fill='x', pady=5)

        self.file_path_entry = tk.Entry(top_frame, width=80)
        self.file_path_entry.pack(side=tk.LEFT, padx=5, pady=5)

        self.file_path_button = tk.Button(top_frame, text='選擇CT2檔案', command=self.select_file)
        self.file_path_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.run_button = tk.Button(top_frame, text='執行', command=self.run_analysis_and_display)
        self.run_button.pack(side=tk.LEFT, padx=5, pady=5)

        # 結果顯示區域
        self.text_result = tk.Text(frame4, width=80, height=20)
        self.text_result.pack(padx=10, pady=10, fill='both', expand=True)
#================================================================ 分頁4 ASM_Yield_page Function================================================================#
    def select_file(self):
        self.CT2_file_path = filedialog.askopenfilename(title="選擇檔案", filetypes=(("CSV Files", ".csv"), ("All Files", ".*")))
        self.file_path_entry.delete("0", tk.END)
        self.file_path_entry.insert(tk.END, self.CT2_file_path)
        return self.CT2_file_path

    def run_analysis_and_display(self):
        if self.CT2_file_path:
            df = pd.read_csv(self.CT2_file_path)
        else:
            return
        
        # New df
        new_df = df[[df.columns[15], df.columns[25]]].rename(columns={
            df.columns[15]: "chip",
            df.columns[25]: "defect"
        })

        """Show"""
        sheetID = df.iloc[:,14].unique()[0]
        status = df.iloc[:,1].unique()[0]
        # 總chip數
        totle_chips = new_df["chip"].unique()
        # 篩選AB16
        defect_code_ab16 = new_df[new_df["defect"]=="AB16"]
        # Ab16 Chip數
        AB16_chips = defect_code_ab16["chip"].unique()
        #篩選AB20
        defect_code_ab20 = new_df[new_df["defect"]=="AB20"]
        # Ab20 Chip數
        AB20_chips = defect_code_ab20["chip"].unique()
        # 排版
        pivot_df = defect_code_ab16.pivot_table(index="chip", values="defect", aggfunc='count').sort_values(by="defect", ascending=False)
        # 美化pivot
        pivot_str = tabulate(pivot_df, headers='keys', tablefmt='grid')
        """Show"""

        z_chip = []
        zp_chip = []
        for chip in AB16_chips:
            defect_code_ab16_counts = defect_code_ab16.loc[(new_df["defect"]=="AB16") & (new_df["chip"]==chip)]
            count_ab16 = len(defect_code_ab16_counts)
            print(defect_code_ab16_counts)

            if count_ab16 <=10:
                z_chip.append(chip)
            if count_ab16 <=50:
                zp_chip.append(chip)

        z_yield = round((len(z_chip)/len(totle_chips))*100, 2)
        zp_yield = round((len(zp_chip)/len(totle_chips))*100, 2)

        # 準備結果字串
        result_str = (
            f"Sheet ID: {sheetID}, 站點: {status}\n"
            f"共計: {len(totle_chips)} chips\n"
            f"Defect AB20: {AB20_chips} \n"
            f"Z規: {len(z_chip)} chips, Yield: {z_yield}%\n"
            f"ZP規: {len(zp_chip)} chips, Yield: {zp_yield}%\n"
            "\n各 Chip defect數量:\n"
            + pivot_str
        )

        # 顯示到 Text widget
        self.text_result.delete("1.0", tk.END)
        self.text_result.insert(tk.END, result_str)
#================================================================ 分頁4 PN_Yield_page ================================================================#

    def create_PN_Yield_page(self):
        thr_frame = tk.Frame(self.page_PN_Yield)
        thr_frame.pack(padx=10, pady=10, fill='both', expand=True)

        # 路徑Entry
        self.PN_dir_entry = tk.Entry(thr_frame, width=60)
        self.PN_dir_entry.pack()
        self.dir_button = tk.Button(thr_frame, text="選擇資料夾路徑", command=self.select_PN_directory)
        self.dir_button.pack()
        
        self.PN_file_list = tk.Listbox(thr_frame, width=80, height=15)
        self.PN_file_list.pack(padx=10, pady=10)

        # 新增：合併並存檔按鈕
        self.btn_merge_save = tk.Button(thr_frame, text="合併", command=self.merge_and_save)
        self.btn_merge_save.pack(pady=5)

        # 新增：Treeview顯示合併資料
        self.merge_tree = ttk.Treeview(thr_frame, columns=[], show='headings')
        self.merge_tree.pack(padx=10, pady=10, fill='both', expand=True)
        # 初始沒有資料，等待合併後設定欄位

        # 文字框
        self.analysis_PN_text = scrolledtext.ScrolledText(thr_frame, width=80, height=10)
        self.analysis_PN_text.pack(padx=10, pady=10)

        #執行按鈕
        run_button = tk.Button(thr_frame, text="Yield分析", command=self.run_PN_analysis)
        run_button.pack(pady=5)
#================================================================ 分頁4 PN_Yield_page Function================================================================#
    def select_PN_directory(self):
        self.directory = filedialog.askdirectory()
        self.PN_csv_file = os.listdir(self.directory)
        if self.directory:
            self.PN_dir_entry.config(state="normal") # 更改為可寫入
            self.PN_dir_entry.delete(0, tk.END) # 清空 Entry
            self.PN_dir_entry.insert(0, self.directory) # 在 Entry 中顯示選擇的路徑
            self.PN_dir_entry.config(text=self.directory) # 顯示
            self.PN_dir_entry.config(state="disabled") # 更改為不可寫入
            self.PN_file_list.delete(0, tk.END)
            for csv in self.PN_csv_file:
                self.PN_file_list.insert(tk.END, csv)

    def merge_and_save(self):
        # 確認資料夾路徑
        dir_path = self.directory
        if not dir_path:
            messagebox.showwarning("警告", "請先選擇資料夾路徑")
            return
        # 取得所有符合條件的CSV檔（此範例假設目前已在file_list）
        files = [os.path.join(dir_path, f) for f in self.PN_file_list.get(0, tk.END) if f.endswith('_sum.csv')]
        if not files:
            messagebox.showwarning("警告", "未找到符合條件的CSV檔案")
            return

        df_list = []
        invalid_RDO1 = False
        for file in files:
            try:
                df = pd.read_csv(file)
            except:
                continue

            df["#Sheet ID"] = df["#Sheet ID"].astype(str)
            df["OPID"] = df["OPID"].astype(str)
            df["Defect Count"] = df["Defect Count"].astype(int)
            df["LED Type"] = df["LED Type"].astype(str)
            df["Target Area"] = df["Target Area"].astype(str)
            df.fillna(0, inplace=True)
            # 刪除 LED Count=0
            df.drop(df[df["LED Count"]==0].index, inplace=True)
            if df.empty:
                continue

            # 檢查OPID
            invalid_opids = df.loc[df["OPID"] != "C2-RDO1", ["#Sheet ID", "OPID"]]
            if not invalid_opids.empty:
                invalid_RDO1 = True
                messagebox.showerror("OPID錯誤", f"{invalid_opids['#Sheet ID'].tolist()} OPID: {invalid_opids['OPID'].tolist()}")
                return
            else:
                df_list.append(df)

        if invalid_RDO1:
            return

        if not df_list:
            messagebox.showwarning("警告", "沒有有效的資料合併")
            return

        # 進行合併
        self.merged_df = pd.concat(df_list, ignore_index=True)

        # 顯示於Treeview
        self.display_merge_dataframe(self.merged_df.loc[:, ["#Sheet ID", "Target Area", "LED Count", "LED Type", "Defect Count", "AB02", "AB06", "Judge"]])


    def run_PN_analysis(self):
        for col in ["#Sheet ID", "OPID", "LED Type", "Target Area"]:
            self.merged_df[col] = self.merged_df[col].astype(str)
            self.merged_df["Defect Count"] = self.merged_df["Defect Count"].astype(int)

        total_W = len(self.merged_df[self.merged_df["LED Type"]=="A"])
        total_O = len(self.merged_df[(self.merged_df["Defect Count"] <=15) & (self.merged_df["Judge"]=="O")])
        total_X = len(self.merged_df[(self.merged_df["Defect Count"] >15) & (self.merged_df["Judge"]=="X")])
        if total_W > 0:
            PN_Yield = round((total_O / total_W) * 100, 2)
        else:
            PN_Yield = 0

        # 將結果顯示在文字框中
        self.analysis_PN_text.delete("1.0", tk.END)
        self.analysis_PN_text.insert(tk.END, 
            f"總Watchs: {total_W}\nO品Watch數量: {total_O}\nX品Watch數量: {total_X}\nYield: {PN_Yield}%"
        )

    def display_merge_dataframe(self, df):
        # 清空
        for col in self.merge_tree.get_children():
            self.merge_tree.delete(col)
        self.merge_tree["columns"] = list(df.columns)
        # 設定標題
        for col in df.columns:
            self.merge_tree.heading(col, text=col)
            self.merge_tree.column(col, width=100)
        # 插入資料
        for _, row in df.iterrows():
            self.merge_tree.insert("", "end", values=list(row))
     
# 創建主窗口
root = tk.Tk()
csv_processor = generate_defect_summary_report(root)

# 運行主循環
root.mainloop()

