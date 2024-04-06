import customtkinter as ctk
from tkinter import filedialog
import win32api
import win32print
import os
import openpyxl as oxl
#エラー発生時の復帰処理
import psutil
#pdf変換
import win32com
import win32com.client
import os
import pythoncom

FONT_TYPE = 'meiryo'

class MyTabView(ctk.CTkTabview):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # create tabs
        self.add("Export")
        self.add("tab 2")

        self._segmented_button.configure(font=(FONT_TYPE, -15))

        # add widgets on tabs
        self.tab1_content = tab_export(master=self.tab("Export"))
        self.tab1_content.grid(row=0, column=0, padx=20, pady=10)

        self.tab2_content = Tab2Content(master=self.tab("tab 2"))
        self.tab2_content.grid(row=0, column=0, padx=20, pady=10)

class tab_export(ctk.CTkFrame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.back_instance = Back_end()
        window_width = self.winfo_screenwidth() // 4.4
        window_height = self.winfo_screenheight() // 5
        self.label = ctk.CTkLabel(master=self, text="", font=(FONT_TYPE, 16))
        self.label.grid(ipadx=window_width, ipady=window_height)
        self.widget()
    
    def widget(self):
        self.read_file_frame = ReadFileFrame(master=self, header_name="ファイル読み込み")
        self.read_file_frame.grid(row=0, column=0, padx=20, pady=(0,400), sticky="ew")
        self.explain_pdf = ctk.CTkLabel(master=self, text='シートを全て.pdfに変換', font=(FONT_TYPE,16))
        self.explain_pdf.grid(row=0, column=0)
        self.Button_pdf = ctk.CTkButton(master=self, text='PDF変換', command=lambda: self.back_instance.pdf_exporter(file_data=self.read_file_frame.get_file_path(), outPutFolder="FOLDERPATH"))
        self.Button_pdf.grid(row=0, column=0, padx=(400,0), pady=(0,0))
        self.explain_png = ctk.CTkLabel(master=self, text='シートを全て.pngに変換', font=(FONT_TYPE,16))
        self.explain_png.grid(row=0, column=0, pady=(100,0))
        self.Button_png = ctk.CTkButton(master=self, text = 'PNG変換', command=lambda:self.pdf_exporter())
        self.Button_png.grid(row=0, column=0, padx=(400,0), pady=(100,0))


class Tab2Content(ctk.CTkFrame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        window_width = self.winfo_screenwidth() // 4.4
        window_height = self.winfo_screenheight() // 5
        self.label = ctk.CTkLabel(master=self, text="", font=(FONT_TYPE, 16))
        self.label.grid(ipadx=window_width, ipady=window_height)
        self.csv_filepath = None
        self.widget()

    def widget(self):
        self.read_file_frame = ReadFileFrame(master=self, header_name="ファイル読み込み")
        self.read_file_frame.grid(row=0, column=0, padx=20, pady=(0,400), sticky="ew")

class ReadFileFrame(ctk.CTkFrame):
    def __init__(self, *args, header_name="ReadFileFrame", **kwargs):
        super().__init__(*args, **kwargs)
        
        self.fonts = (FONT_TYPE, 15)
        self.header_name = header_name

        # フォームのセットアップをする
        self.setup_form()

    def setup_form(self):
        # 行方向のマスのレイアウトを設定する。リサイズしたときに一緒に拡大したい行をweight 1に設定。
        self.grid_rowconfigure(0, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)

        # フレームのラベルを表示
        self.label = ctk.CTkLabel(self, text=self.header_name, font=(FONT_TYPE, 11))
        self.label.grid(row=0, column=0, padx=20, sticky="w")

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox = ctk.CTkEntry(master=self, placeholder_text="xlsx ファイルを読み込む", width=120, font=self.fonts)
        self.textbox.grid(row=1, column=0, padx=10, pady=(0,10), sticky="ew")

        # ファイル選択ボタン
        self.button_select = ctk.CTkButton(master=self, 
            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),   # ボタンを白抜きにする
            command=self.button_select_callback, text="ファイル選択", font=self.fonts)
        self.button_select.grid(row=1, column=1, padx=10, pady=(0,10))
        
        # 開くボタン
        self.button_open = ctk.CTkButton(master=self, command=self.button_open_callback, text="開く", font=self.fonts)
        self.button_open.grid(row=1, column=2, padx=10, pady=(0,10))

    def button_select_callback(self):
        """
        選択ボタンが押されたときのコールバック。ファイル選択ダイアログを表示する
        """
        # エクスプローラーを表示してファイルを選択する
        file_name = ReadFileFrame.file_read()

        if file_name is not None:
            # ファイルパスをテキストボックスに記入
            self.textbox.delete(0, ctk.END)
            self.textbox.insert(0, file_name)

    def button_open_callback(self):
        """
        開くボタンが押されたときのコールバック。暫定機能として、ファイルの中身をprintする
        """
        try:
            file_name = self.textbox.get()
            if file_name is not None or len(file_name) != 0:
                with open(file_name) as f:
                    data = f.read()
                    print(data)
        except:
            None
            
    @staticmethod
    def file_read():
        """
        ファイル選択ダイアログを表示する
        """
        current_dir = os.path.abspath(os.path.dirname(__file__))
        file_path = filedialog.askopenfilename(filetypes=[("xlsxファイル","*.xlsx")],initialdir=current_dir)

        if len(file_path) != 0:
            return file_path
        else:
            # ファイル選択がキャンセルされた場合
            return None

    def get_file_path(self):
        """
        テキストボックスからファイルパスを取得して返す。
        """
        return self.textbox.get()

class Back_end():
    def __init__(self):
        pass

    def printer(self, file_path):
        try:
            # Use win32api to print the file directly
            win32api.ShellExecute(
                0,
                "print",
                file_path,
                '/c:"%s"' % win32print.GetDefaultPrinter(),
                ".",
                0
            )
        except:
            print("エラー : 印刷エラーです。コピー機を確認してください")
    
    def pdf_exporter(self, file_data, outPutFolder):
        try:
            if file_data == True:
                pythoncom.CoInitialize()
                excel = win32com.client.Dispatch("Excel.Application")
                if not os.path.exists(outPutFolder):
                    os.makedirs(outPutFolder)
                excel_data = oxl.load_workbook(file_data)
                sheets_name = excel_data.sheetnames

                base, ext = os.path.splitext(file_data)
                if(ext == '.xlsx' and '~$' not in base):
                    for j in sheets_name:
                        wb1 = excel.Workbooks.Open(file_data)
                        wb1.WorkSheets(j).Select()
                        output_path = os.path.join(outPutFolder, f'{j}.pdf')
                        if os.path.exists(output_path):
                            os.remove(output_path)
                        wb1.ActiveSheet.ExportAsFixedFormat(0, output_path)
                    wb1.Close()
        except Exception:
            self.task_kill()

    #Excelが既に開かれていた場合にTask Killを実行
    def task_kill(self):
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'EXCEL.EXE':
                pid = proc.info['pid']
                os.kill(pid, 9)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("EXCEL TOOL")
        self.geometry("960x560")

        self.tab_view = MyTabView(master=self)
        self.tab_view.grid(row=0, column=0, padx=15)

if __name__ == '__main__':
    app = App()
    app.mainloop()
