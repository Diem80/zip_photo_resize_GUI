import zipfile
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, scrolledtext
from tkinter.constants import DISABLED, NORMAL, END
from tkinter.messagebox import showerror
import os, sys, shutil, subprocess
import configparser
from utils import *
from PIL import Image
from io import BytesIO

class PhotoResize(tk.Tk):
    def __init__(self):
        super().__init__()

        # determine if the application is a frozen `.exe` (e.g. pyinstaller --onefile) 
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            application_path = os.path.dirname(__file__)
        config_name = 'config.ini'
        self.config_path = os.path.join(application_path, config_name)

        # 설정파일 읽기
        config = configparser.ConfigParser()
        config.read(self.config_path, encoding='utf-8') 

        self.recentPath = config['PATH']['recentPath']
        self.destinationPath = config['PATH']['destinationPath']

        self.running = False

        # configure the root window
        self.title('Photo ZIP File Resize')
        self.resizable(False, False)
        # self.geometry("640x480+300+100")

        self.style_1 = ttk.Style(self)
        self.style_1.layout('text.Horizontal.TProgressbar1',
                    [('Horizontal.Progressbar.trough',
                        {'children': [('Horizontal.Progressbar.pbar',
                                        {'side': 'left', 'sticky': 'ns'})],
                        'sticky': 'nswe'}),
                        ('Horizontal.Progressbar.label', {'sticky': ''})])

        self.style_2 = ttk.Style(self)
        self.style_2.layout('text.Horizontal.TProgressbar2',
                    [('Horizontal.Progressbar.trough',
                        {'children': [('Horizontal.Progressbar.pbar',
                                        {'side': 'left', 'sticky': 'ns'})],
                        'sticky': 'nswe'}),
                        ('Horizontal.Progressbar.label', {'sticky': ''})])

        # 파일 프레임
        frm_file = tk.LabelFrame(self, text="파일 선택", borderwidt=0)
        frm_file.pack(fill="x", padx=5, pady=5, ipady=5)

        # 파일 리스트
        self.list_box = tk.Listbox(frm_file, selectmode = tk.EXTENDED, width="80")
        self.list_box.pack(side="left")
        
        # 파일 버튼 프레임
        frame_button = tk.Frame(frm_file)
        frame_button.pack(side="right", padx=5, pady=5, ipady=5)

        self.btn_add = tk.Button(frame_button, text="파일 추가", command=self.add_files)
        self.btn_add.pack(padx=5, pady=5, ipadx=10)
        self.btn_remove = tk.Button(frame_button, text="파일 삭제", command=self.remove_files)
        self.btn_remove.pack(padx=5, pady=5, ipadx=10)
        self.btn_remove_all = tk.Button(frame_button, text="전체 삭제", command=self.remove_all)
        self.btn_remove_all.pack(padx=5, pady=5, ipadx=10)

        # 저장 경로 프레임
        frame_path = tk.LabelFrame(self, text="저장경로")
        frame_path.pack(fill="x", padx=5, pady=5, ipady=5)

        self.txt_dest_path = tk.Entry(frame_path)
        self.txt_dest_path.insert(0, self.destinationPath)
        self.txt_dest_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4)

        self.btn_dest_path = tk.Button(frame_path, text="찾아보기", width=10, command=self.browse_dest_path)
        self.btn_dest_path.pack(side="left", padx=5, pady=5)

        btn_dest_open = tk.Button(frame_path, text="열기", width=10, command=self.browse_dest_open)
        btn_dest_open.pack(side="right", padx=5, pady=5)

        # 실행 & 종료 & 정지 버튼
        frame_run = tk.Frame(self)
        frame_run.pack(fill="x", padx=5, pady=5, ipady=5)

        self.btn_stop = tk.Button(frame_run, text="중지", width=20, command=self.stop, state=DISABLED)
        self.btn_stop.pack(side="left", padx=5, pady=5)

        btn_quit = tk.Button(frame_run, text="종료", width=20, command=self.quit)
        btn_quit.pack(side="right", padx=5, pady=5)

        self.btn_run = tk.Button(frame_run, text="실행", width=20, command=self.run)
        self.btn_run.pack(side="right", padx=5, pady=5)

        # 결과
        frame_result = tk.LabelFrame(self, text="결과")
        frame_result.pack(fill="x", padx=5, pady=5, ipady=5)

        self.txt_result = scrolledtext.ScrolledText(frame_result)
        self.txt_result.pack(fill="x", expand=True)

        # 프로그레스바
        self.progress_file = ttk.Progressbar(self, orient='horizontal', mode='determinate', style="text.Horizontal.TProgressbar1")
        self.progress_file.pack(fill="x")

        self.progress_total = ttk.Progressbar(self, orient='horizontal', mode='determinate', style="text.Horizontal.TProgressbar2")
        self.progress_total.pack(fill="x")
    
    # 설정파일 저장
    def config_save(self):
        config = configparser.ConfigParser()

        config['PATH'] = {}
        config['PATH']['recentPath'] = os.path.realpath(self.recentPath)
        config['PATH']['destinationPath'] = os.path.realpath(self.destinationPath)

        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            config.write(configfile)

    # 리사이즈
    def zip_photo_resize(self, file_path, dest_path="./", size=1280):
        RESIZE = size
        file = os.path.basename(file_path)
        filename, ext = os.path.splitext(file)

        if not ext != "zip":
            # print("It's not a .zip file")
            return False

        tempdir = os.path.join(dest_path, "temp")
        unzip_path = os.path.join(tempdir, filename)

        createFolder(tempdir)
        createFolder(unzip_path)

        with zipfile.ZipFile(file_path, 'r') as ori_zip:
            i = 0
            dest_name = os.path.join(dest_path, file)

            with zipfile.ZipFile(dest_name, "w") as new_zip:
                self.progress_file['maximum'] = int(len(ori_zip.infolist()))
                cnt_dir = len([info for info in ori_zip.infolist() if info.is_dir()])

                for info in ori_zip.infolist():
                    if info.is_dir():
                        continue
                    
                    try:
                        info.filename = info.filename.encode('cp437').decode('euc-kr').replace(filename + "/", "")
                    except:
                        info.filename = info.filename.encode('utf-8').decode('euc-kr').replace(filename + "/", "")

                    try:
                        data = ori_zip.open(info)
                        with Image.open(data) as ori_image:
                            width, height = ori_image.size
                            exif = ori_image.getexif()
                            format = ori_image.format

                            if width > height:
                                new_photo = ori_image.resize((RESIZE, int(RESIZE * height / width)), Image.Resampling.LANCZOS)
                            else:
                                new_photo = ori_image.resize((int(RESIZE * width / height), RESIZE), Image.Resampling.LANCZOS)
                            image_file = BytesIO()
                            new_photo.save(image_file, format=format, quality=95, exif=exif)
                            new_zip.writestr(info.filename if cnt_dir > 1 else os.path.basename(info.filename), image_file.getvalue(), compress_type=zipfile.ZIP_DEFLATED)
                            i += 1
                    except Exception as e:
                        ori_zip.extract(info, unzip_path)

                    self.progress_file['value'] = i
                    self.style_1.configure("text.Horizontal.TProgressbar1", text=f"{info.filename} ({int(self.progress_file['value'])}/{self.progress_file['maximum']})")
                    self.progress_file.update()

            self.style_1.configure("text.Horizontal.TProgressbar1", text="Finish")
            new_name = os.path.join(dest_path, f"{filename} [{i}장]{ext}")
            uniq_rename(dest_name, new_name)
            
        if len(os.listdir(unzip_path)) == 0:
            shutil.rmtree(unzip_path, ignore_errors=False, onerror=handleRemoveReadonly)

    # 파일 추가
    def add_files(self):
        filetypes = [('Zip files', '.zip')]
        filenames = filedialog.askopenfilenames(title='Open a file', filetypes=filetypes, initialdir=self.recentPath)
        
        if filenames == '': # 사용자가 취소를 누를 때
            return

        current_list = self.list_box.get(0, END)

        for file in filenames:
            file_path = os.path.realpath(file)
            if file_path not in current_list:
                self.list_box.insert(END, file_path)
        
        self.recentPath = os.path.dirname(filenames[-1])
        self.config_save()

    # 파일 제거
    def remove_files(self):
        selected = self.list_box.curselection()
  
        for index in selected[::-1]:
            self.list_box.delete(index)

    # 전체 제거
    def remove_all(self):
        self.list_box.delete(0, END)

    # 저장 경로 (폴더)
    def browse_dest_path(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected == '': # 사용자가 취소를 누를 때
            return
        self.txt_dest_path.delete(0, END)
        self.txt_dest_path.insert(0, os.path.realpath(folder_selected))

        self.destinationPath = folder_selected
        self.config_save()

    # 저장 경로 열기
    def browse_dest_open(self):
        path = self.txt_dest_path.get()
        subprocess.run(['explorer', os.path.realpath(path)])
    
    # 변환 실행
    def run(self):
        self.running = True
        self.txt_result.delete("1.0", END)
        file_list = self.list_box.get(0, END)
        dest_path = self.txt_dest_path.get()

        if len(file_list) == 0:
            showerror("Warring", "No files")

        self.btn_state(DISABLED)

        self.progress_total['value'] = 0

        self.progress_total['maximum'] = int(len(file_list))
        for file_path in file_list:
            self.print_msg(file_path)

            self.progress_total['value'] += 1
            self.style_2.configure("text.Horizontal.TProgressbar2", text=f"{os.path.basename(file_path)} ({int(self.progress_total['value'])}/{self.progress_total['maximum']})")
            self.progress_total.update()
            
            self.zip_photo_resize(file_path, dest_path)
            if not self.running:
                break

        self.style_2.configure("text.Horizontal.TProgressbar2", text="Finish")
        self.print_msg("완료")
        
        self.btn_state(NORMAL)

    # 종료
    def quit(self):
        self.config_save()
        self.destroy()

    # 중지
    def stop(self):
        self.running = False
        self.print_msg("중지")
        self.btn_stop.configure(state=DISABLED)

    # 메시지 출력
    def print_msg(self, msg):
        self.txt_result.insert(END, msg + '\n')

    # 실행 중 버튼 상태
    def btn_state(self, state):
        self.btn_add.configure(state=state)
        self.btn_remove.configure(state=state)
        self.btn_remove_all.configure(state=state)
        self.btn_dest_path.configure(state=state)
        self.btn_run.configure(state=state)

        if state == DISABLED:
            self.btn_stop.configure(state=NORMAL)
        else:
            self.btn_stop.configure(state=DISABLED)

def main():
    app = PhotoResize()
    app.mainloop()

if __name__ == '__main__':
    main()