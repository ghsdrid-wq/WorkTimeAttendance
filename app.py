import os
import threading
import time
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog, messagebox, font
from datetime import datetime
import worker
import configparser
import shutil
import pandas as pd
from idlelib.tooltip import Hovertip

def load_embedded_font(root):
    font_path = os.path.join(os.path.dirname(__file__), "fonts", "IBMPlexSansThai-Regular.ttf")
    try:
        root.tk.call('font', 'create', 'thaiFont', '-family', 'IBM Plex Sans Thai', '-size', '10', '-weight', 'normal')
        root.tk.call('font', 'configure', 'thaiFont', '-file', font_path)
    except Exception as e:
        print("Font load error:", e)

def set_default_font():
    for f in ["TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"]:
        try:
            font.nametofont(f).configure(family="IBM Plex Sans Thai", size=10)
        except:
            pass
       
class WorkTimeProcessor:
    def __init__(self, root: ttk.Window):
        self.root = root
        self.root.title("Work Time Attendance")
        self.root.resizable(False, False)
        self.running = False
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="IBM Plex Sans Thai", size=10)
        self.shift_base = ttk.StringVar(value="16:00")
        self.shift_offset = ttk.IntVar(value=3)

        self.shift_base_night = ttk.StringVar(value="00:00")
        self.shift_offset_night = ttk.IntVar(value=3)

        self.include_data = ttk.BooleanVar(value=False)

        self.keep_unpaired = ttk.BooleanVar(value=False)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(5, weight=1)
        
        self.load_config()
        self.create_ui()

     
    def load_config(self):
        self.config = configparser.ConfigParser()
        self.config_file = "config.ini"

        # ถ้ายังไม่มี config.ini ให้สร้างใหม่
        if not os.path.exists(self.config_file):
            self.config['PATH'] = {
                'last_import': '',
                'last_export': '',
                'last_import_dir': '',
                'last_export_dir': ''
            }
            self.config['SETTINGS'] = {
                'base_day': '16:00',
                'offset_day': '3',
                'base_night': '00:00',
                'offset_night': '3',
                'keep_unpaired': 'False',
                'include_data': 'False'
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)

        # โหลดไฟล์
        self.config.read(self.config_file, encoding='utf-8')

        # โหลดค่าจาก SETTINGS
        self.shift_base.set(self.config.get("SETTINGS", "base_day", fallback="16:00"))
        self.shift_offset.set(self.config.getint("SETTINGS", "offset_day", fallback=3))
        self.shift_base_night.set(self.config.get("SETTINGS", "base_night", fallback="00:00"))
        self.shift_offset_night.set(self.config.getint("SETTINGS", "offset_night", fallback=3))
        self.keep_unpaired.set(self.config.getboolean("SETTINGS", "keep_unpaired", fallback=False))
        self.include_data.set(self.config.getboolean("SETTINGS", "include_data", fallback=False))


    def save_config(self):
        if "SETTINGS" not in self.config:
            self.config["SETTINGS"] = {}
        self.config.set("SETTINGS", "base_day", self.shift_base.get())
        self.config.set("SETTINGS", "offset_day", str(self.shift_offset.get()))
        self.config.set("SETTINGS", "base_night", self.shift_base_night.get())
        self.config.set("SETTINGS", "offset_night", str(self.shift_offset_night.get()))
        self.config.set("SETTINGS", "keep_unpaired", str(self.keep_unpaired.get()))
        self.config.set("SETTINGS", "include_data", str(self.include_data.get()))
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def create_ui(self):
        pad = 8

        # === เพิ่ม Notebook (Tabbar) ===
        '''self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=7, column=0, columnspan=2, padx=8, pady=8, sticky="nsew")

        # === เพิ่ม Tab สำหรับ Setting/Base Shift ===
        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_settings, text="Settings")'''
        # === Frame รวม Settings ===
        frm_settings = ttk.Frame(self.root)
        frm_settings.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        # ให้ root และ settings รองรับการขยาย
        self.root.columnconfigure(0, weight=1)
        frm_settings.columnconfigure(0, weight=1)   # Export
        frm_settings.columnconfigure(1, weight=0)   # Base Shift

        ttk.Label(self.root, text="Input Excel File:", bootstyle="secondary").grid(row=0, column=0, padx=pad, pady=(pad,0), sticky=W)
        
        self.txt_input = ttk.Entry(self.root, width=80)
        self.txt_input.grid(row=1, column=0, padx=pad, pady=0, sticky=EW)
        last_in = self.config.get("PATH", "last_import", fallback="")

        if last_in and os.path.exists(last_in):
            self.txt_input.insert(0, last_in)

        self.btn_browse = ttk.Button(self.root, text="Browse", bootstyle=PRIMARY, command=self.browse)
        self.btn_browse.grid(row=1, column=1, padx=(0, pad), pady=0, sticky=EW)

        ttk.Label(self.root, text="Progress:", bootstyle="secondary").grid(row=2, column=0, padx=pad, pady=(pad,0), sticky=W)
        self.progress = ttk.Progressbar(self.root, maximum=100, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=2, padx=pad, pady=0, sticky=EW)

        ttk.Label(self.root, text="Log:", bootstyle="secondary").grid(row=4, column=0, padx=pad, pady=(pad,0), sticky=W)
        self.txt_log = ScrolledText(self.root, height=15, width=100)
        self.txt_log.grid(row=5, column=0, columnspan=2, padx=pad, pady=0, sticky=EW)

        # ===== Export Labelframe =====
        export_setting = ttk.Labelframe(frm_settings, text="Export", bootstyle="secondary")
        export_setting.grid(row=0, column=0, padx=8, pady=8, sticky="nsew")
        export_setting.columnconfigure(0, weight=1)

        # ===== Base Shift Labelframe =====
        frm_shift = ttk.Labelframe(frm_settings, text="Base Shift")
        frm_shift.grid(row=0, column=1, padx=8, pady=8, sticky="nsew")
        frm_shift.configure(bootstyle="secondary")
        Hovertip(frm_shift, "ใส่เวลาเริ่มของกะในรูปแบบ HH:MM เช่น 16:00 ใส่ชั่วโมง ± เป็นตัวเลข เช่น 3\n" \
                            "ตัวอย่าง : เวลาฐานที่ตั้งไว้คือ 16:00(กะบ่าย) ± 3 ชั่วโมง จะได้เป็น 13:00-19:00\n" \
                            "หากมีการเข้าทีงานในช่วงเวลา 13:00-19:00 จะถือว่าเป็นการเข้าทำงานของ กะบ่าย\n" \
                            "หมายเหตุ : ชั่วโมง ± คือช่วงเวลาคลาดเคลื่อน มีไว้เพื่อป้องกันกรณีที่วันไหนมีการเริ่มงานก่อนหรือหลังเวลาฐานกะ\n", hover_delay=500)
        # ให้ element ภายในจัด layout ดีขึ้น
        frm_shift.columnconfigure(0, weight=0)
        frm_shift.columnconfigure(1, weight=0)
        frm_shift.columnconfigure(2, weight=0)
        frm_shift.columnconfigure(3, weight=0)
        frm_shift.columnconfigure(4, weight=0)

        export_setting.columnconfigure(0, weight=1)
        # -------- Row 0 : กะบ่าย --------
        ttk.Label(frm_shift, text="กะบ่าย").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frm_shift, textvariable=self.shift_base, width=6).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frm_shift, text="±").grid(row=0, column=2, padx=5, pady=5)
        ttk.Entry(frm_shift, textvariable=self.shift_offset, width=3).grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(frm_shift, text="ชั่วโมง").grid(row=0, column=4, padx=5, pady=5)

        # -------- Row 1 : กะดึก --------
        ttk.Label(frm_shift, text="กะดึก").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frm_shift, textvariable=self.shift_base_night, width=6).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frm_shift, text="±").grid(row=1, column=2, padx=5, pady=5)
        ttk.Entry(frm_shift, textvariable=self.shift_offset_night, width=3).grid(row=1, column=3, padx=5, pady=5)
        ttk.Label(frm_shift, text="ชั่วโมง").grid(row=1, column=4, padx=5, pady=5)

        ttk.Checkbutton(
            export_setting, 
            text="Export DATA sheet", 
            variable=self.include_data, 
            bootstyle="info"
        ).grid(row=1, column=0, padx=pad, pady=(pad,0), sticky=W)
        
        ttk.Checkbutton(
            export_setting,
            text="เก็บรายชื่อ IN/OUT ที่ไม่สมบูรณ์",
            variable=self.keep_unpaired,
            bootstyle="info"
        ).grid(row=2, column=0,padx=pad, pady=(pad,0), sticky=W)


        self.btn_process = ttk.Button(self.root, text="► PROCESS", bootstyle=SUCCESS, command=self.start)
        self.btn_process.grid(row=8, column=0, columnspan=2, padx=pad, pady=(pad,pad), sticky=EW)

    def browse(self):
        last_dir = self.config.get("PATH", "last_import_dir", fallback="")

        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xls *.xlsx")],
            initialdir=last_dir if os.path.isdir(last_dir) else None
        )

        if path:
            self.txt_input.delete(0, "end")
            self.txt_input.insert(0, path)

            # บันทึกไฟล์ + โฟลเดอร์
            self.config.set("PATH", "last_import", path)
            self.config.set("PATH", "last_import_dir", os.path.dirname(path))
            self.save_config()

    def start(self):
        input_file = self.txt_input.get().strip()
        if not os.path.exists(input_file):
            self.log("❌ คุณยังไม่ได้เลือกไฟล์!")
            messagebox.showwarning(
                "ไม่มีไฟล์ที่เลือก",
                "กรุณาเลือกไฟล์ Excel ก่อนทำงาน"
            )
            return

        out_folder = os.path.dirname(input_file)
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_file = os.path.join(out_folder, f"output_{now}.xlsx")

        self.running = True
        self.progress['value'] = 0
        self.log("▶ เริ่มประมวลผล...\n")

        thread = threading.Thread(
            target=self.do_work,
            args=(
                input_file,
                self.shift_base.get(),         # "16:00"
                self.shift_offset.get(),       # 3
                self.shift_base_night.get(),   # "00:00"
                self.shift_offset_night.get()  # 3
            ),
            daemon=True
        )  
        thread.start()
        self.btn_process.configure(state=DISABLED)
        self.btn_browse.configure(state=DISABLED)

    def do_work(self, input_file, base_day, offset_day, base_night, offset_night):
        try:
            # Step 1
            self.log("📌 ขั้นตอนที่ 1: โหลดข้อมูลจากไฟล์")
            self.progress['value'] = 10
            self.update()

            df = pd.read_excel(input_file)

            # Step 2
            self.log("📌 ขั้นตอนที่ 2: จับคู่ข้อมูล เข้า/ออก (IN/OUT)")
            self.progress['value'] = 40
            self.update()

            # (simulate heavy pairing just for display)
            total = len(df)
            for i in range(0, total, max(1, total//5)):
                self.log(f"    → ประมวลผลแถวที่ {i}/{total}")
                self.update()
                time.sleep(0.1)

            # Step 3
            self.log("📌 ขั้นตอนที่ 3: คำนวณชั่วโมงทำงาน")
            self.progress['value'] = 70
            self.update()
            time.sleep(0.3)

            # Step 4: Real Processing
            worker.process(
                input_file,
                self.output_file,
                include_data=self.include_data.get(),
                keep_unpaired=self.keep_unpaired.get(),
                base_day=base_day,
                offset_day=offset_day,
                base_night=base_night,
                offset_night=offset_night
            )

            # Step 5 Done
            self.log("📌 ขั้นตอนที่ 4: สร้างไฟล์สรุปผลเรียบร้อย")
            self.progress['value'] = 100
            self.update()

            last_export_dir = self.config.get("PATH", "last_export_dir", fallback="")

            save = filedialog.asksaveasfilename(
                title="Save Output File",
                initialfile=f"summary_worktime.xlsx",
                defaultextension=".xlsx",
                initialdir=last_export_dir if os.path.isdir(last_export_dir) else None,
                filetypes=[("Excel Files", "*.xlsx")]
            )

            if save:
                try:
                    shutil.copy2(self.output_file, save)
                    self.log(f"💾 บันทึกไฟล์สำเร็จ: {save}")
                    messagebox.showinfo(
                        "การบันทึกเสร็จสิ้น",
                        f"บันทึกไฟล์สำเร็จ!\n\nที่อยู่ไฟล์:\n{save}"
                    )
                    self.config.set("PATH", "last_export", save)
                    self.config.set("PATH", "last_export_dir", os.path.dirname(save))
                    self.save_config()

                    status = "saved"

                except Exception as e:
                    self.log(f"❌ บันทึกไฟล์ล้มเหลว: {e}")
                    status = "error"
            else:
                self.log("⚠ การบันทึกถูกยกเลิก")
                status = "cancelled"

            # ลบไฟล์ชั่วคราวเสมอ
            if os.path.exists(self.output_file):
                try:
                    os.remove(self.output_file)
                except Exception as e:
                    self.log(f"⚠ ไม่สามารถลบไฟล์ชั่วคราวได้: {e}")

            # แสดงสถานะสรุป
            if status == "saved":
                self.log("\n🎉 เสร็จสิ้น!")
            elif status == "cancelled":
                self.log("\n⏹ การทำงานถูกยกเลิกโดยผู้ใช้ (ไม่มีไฟล์ถูกบันทึก)")
            elif status == "success":
                messagebox.showinfo("การบันทึกเสร็จสิ้น",f"ระบบได้บันทึกไฟล์เรียบร้อยแล้ว\n\nที่อยู่ไฟล์:\n{save}")
            elif status == "error":
                messagebox.showerror("เกิดข้อผิดพลาด",f"พบข้อผิดพลาดระหว่างการทำงาน:\n\n{e}")
            else:
                self.log("\n❌ มีข้อผิดพลาดเกิดขึ้น")

            # <<< ลบไฟล์ชั่วคราวเสมอ
            if os.path.exists(self.output_file):
                try:
                    os.remove(self.output_file)
                except Exception as e:
                    self.log(f"⚠ ไม่สามารถลบไฟล์ชั่วคราวได้: {e}")

        except Exception as e:
            self.log(f"❌ ผิดพลาด: {e}")
        finally:
            self.running = False
            self.btn_process.configure(state=NORMAL)
            self.btn_browse.configure(state=NORMAL)


    def log(self, msg):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")

    def update(self):
        self.root.update_idletasks()
        self.root.update()


if __name__ == "__main__":
    app = ttk.Window(themename="flatly")

    # โหลดฟอนต์จากไฟล์ในโฟลเดอร์ fonts
    load_embedded_font(app)

    # ตั้งฟอนต์ default สำหรับ UI
    set_default_font()

    WorkTimeProcessor(app)
    app.mainloop()

