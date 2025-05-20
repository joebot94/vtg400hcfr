import tkinter as tk
from tkinter import ttk
import serial
import serial.tools.list_ports
import re
import win32gui
from pywinauto import Desktop
from PIL import ImageGrab
import pytesseract

POLL_INTERVAL = 0.5  # HCFR in seconds

class CombinedController:
    def __init__(self, master):
        self.master = master
        master.title("VTG 400 Serial Control with HCFR")

        # Theme colors
        self.bg = '#222222'; self.fg = '#eeeeee'
        self.btn_bg = '#333333'; self.btn_fg = '#eeeeee'
        self.on_bg = '#28a745'; self.off_bg = '#dc3545'
        master.configure(bg=self.bg)

        self.serial_port = None
        self.hcfr_active = False

        # Serial port UI...
        tk.Label(master, text="Serial Port:", bg=self.bg, fg=self.fg)\
          .grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.port_combo = ttk.Combobox(master, values=self.get_serial_ports(), width=15)
        self.port_combo.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(master, text="Connect", command=self.connect_serial,
                  bg=self.btn_bg, fg=self.btn_fg)\
          .grid(row=0, column=2, padx=5, pady=2)

        # Status/readouts...
        self.status_label = tk.Label(master, text="Status: Disconnected",
                                     bg=self.bg, fg=self.fg)
        self.status_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=5)
        self.temp_label = tk.Label(master, text="Temperature: N/A",
                                   bg=self.bg, fg=self.fg)
        self.temp_label.grid(row=2, column=0, columnspan=3, sticky="w", padx=5)
        self.ire_label = tk.Label(master, text="IRE Level: N/A",
                                  bg=self.bg, fg=self.fg)
        self.ire_label.grid(row=3, column=0, columnspan=3, sticky="w", padx=5)

        # Power buttons...
        tk.Button(master, text="Power On", command=lambda: self.send_serial("1P"),
                  bg=self.on_bg, fg='white')\
          .grid(row=4, column=0, sticky="ew", padx=5, pady=5)
        tk.Button(master, text="Power Off", command=lambda: self.send_serial("0P"),
                  bg=self.off_bg, fg='white')\
          .grid(row=4, column=1, sticky="ew", padx=5, pady=5)

        # IRE buttons...
        ire_frame = tk.LabelFrame(master, text="IRE Levels", bg=self.bg, fg=self.fg)
        ire_frame.grid(row=5, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.ire_buttons = {}
        for idx, v in enumerate(range(0, 101, 10)):
            btn = tk.Button(ire_frame, text=f"{v} IRE", bg=self.btn_bg, fg=self.btn_fg,
                            command=lambda vv=v: self.set_ire(vv))
            btn.grid(row=idx//6, column=idx%6, padx=3, pady=3)
            self.ire_buttons[v] = btn

        # Pattern buttons...
        pat_frame = tk.LabelFrame(master, text="Test Patterns", bg=self.bg, fg=self.fg)
        pat_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.patterns = [
            (15,"Window20"), (14,"Window80"), (16,"VarIRE"),
            (17,"FullScreen"), (6,"4x4Cross"), (7,"Coarse"),
            (8,"FineCross"), (13,"ColorBar")
        ]
        self.pattern_buttons = {}
        for idx,(num,name) in enumerate(self.patterns):
            btn = tk.Button(pat_frame, text=name, bg=self.btn_bg, fg=self.btn_fg,
                            command=lambda n=num,b=name: self.select_pattern(n,b))
            btn.grid(row=idx//4, column=idx%4, padx=3, pady=3)
            self.pattern_buttons[name] = btn

        # Color selection...
        colors = [
            ("Black","#000000","0*10#"), ("Blue","#0000FF","1*10#"),
            ("Green","#00FF00","2*10#"), ("Cyan","#00FFFF","3*10#"),
            ("Red","#FF0000","4*10#"), ("Magenta","#FF00FF","5*10#"),
            ("Yellow","#FFFF00","6*10#"), ("White","#FFFFFF","7*10#")
        ]
        self.color_map = {lbl:(cmd,lbl) for lbl,_,cmd in colors}
        self.color_buttons = {}
        # keep original bg/fg for restoration
        self.color_original_colors = {}

        col_frame = tk.LabelFrame(master, text="Color Selection", bg=self.bg, fg=self.fg)
        col_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        for idx,(lbl,bgc,cmd) in enumerate(colors):
            # choose readable text color
            if lbl == "Blue":
                fgc = 'white'
            elif lbl == "Cyan":
                fgc = 'black'
            else:
                fgc = 'white' if bgc in ('#000000','#FF0000','#FF00FF') else 'black'

            # store originals
            self.color_original_colors[lbl] = (bgc, fgc)

            btn = tk.Button(col_frame, text=lbl, bg=bgc, fg=fgc,
                            relief=tk.RAISED, bd=2,
                            command=lambda c=cmd, b=lbl: self.select_color(c,b))
            btn.grid(row=idx//4, column=idx%4, padx=3, pady=3)
            self.color_buttons[lbl] = btn

        # Resolution buttons...
        res_frame = tk.LabelFrame(master, text="Resolution", bg=self.bg, fg=self.fg)
        res_frame.grid(row=8, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.resolutions = [
            ("240p","001*99="), ("NTSC/U","001*07="), ("NTSC/J","002*07="),
            ("PAL","003*07="), ("480p","001*06="), ("576p","002*06="),
            ("720p","004*06="), ("1080i","010*06=")
        ]
        self.res_buttons = {}
        for idx,(lbl,cmd) in enumerate(self.resolutions):
            btn = tk.Button(res_frame, text=lbl, bg=self.btn_bg, fg=self.btn_fg,
                            command=lambda c=cmd,b=lbl: self.select_resolution(c,b))
            btn.grid(row=idx//4, column=idx%4, padx=3, pady=3)
            self.res_buttons[lbl] = btn

        # HCFR toggle...
        self.hcfr_var = tk.IntVar()
        tk.Checkbutton(master, text="Enable HCFR", variable=self.hcfr_var,
                       command=self.toggle_hcfr, bg=self.btn_bg, fg=self.btn_fg,
                       selectcolor=self.bg, activebackground=self.btn_bg,
                       activeforeground=self.btn_fg)\
          .grid(row=9, column=0, columnspan=3, pady=5)

        # Footer credit
        tk.Label(master, text="Made with ChatGPT by Joebot",
                 bg=self.bg, fg=self.fg, font=("Arial",9,"italic"))\
          .grid(row=10, column=0, columnspan=3, pady=(10,5))

        for i in range(3):
            master.grid_columnconfigure(i, weight=1)

        # Reverse-lookups
        self._pattern_map = {num:name for num,name in self.patterns}
        self._res_map     = {cmd.rstrip('='):lbl for lbl,cmd in self.resolutions}
        # case-insensitive HCFR cues → our color names
        self.hcfr_color_map = {
            "red primary":       "Red",
            "green primary":     "Green",
            "blue primary":      "Blue",
            "cyan secondary":    "Cyan",
            "magenta secondary": "Magenta",
            "yellow secondary":  "Yellow",
            "white":             "White",
        }

    # — Serial helpers —

    def get_serial_ports(self):
        return [p.device for p in serial.tools.list_ports.comports()]

    def connect_serial(self):
        port = self.port_combo.get()
        if not port:
            self.status_label.config(text="Select port")
            return
        try:
            self.serial_port = serial.Serial(port, 9600, timeout=1)
            self.serial_port.write(b'N')
            self.master.after(100, self.check_model)
        except Exception as e:
            self.status_label.config(text=f"Error: {e}")

    def check_model(self):
        try:
            resp = self.serial_port.readline().decode().strip()
            mn = ("VTG 400" if "60-564-01" in resp else
                  "VTG 400D" if "60-564-02" in resp else
                  "VTG 400DVI" if "60-564-03" in resp else
                  "Unknown")
            self.status_label.config(text=f"Connected: {mn}")
            # start polling
            self.master.after(500, self.poll_ire)
            self.master.after(500, self.poll_temperature)
            self.master.after(500, self.poll_pattern)
            self.master.after(500, self.poll_resolution)
        except Exception as e:
            self.status_label.config(text=f"Model check error: {e}")

    def send_serial(self, cmd):
        if self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.write(cmd.encode())
            except Exception as e:
                self.status_label.config(text=f"Send error: {e}")
        else:
            self.status_label.config(text="Serial not open")

    # — IRE —

    def set_ire(self, v):
        self.send_serial(f"{v}*15#")
        self.ire_label.config(text=f"IRE Level: {v}")
        self._highlight_ire(v)

    def poll_ire(self):
        if self.serial_port and self.serial_port.is_open:
            self.send_serial('15#')
            self.master.after(100, self.read_ire)
        self.master.after(20000, self.poll_ire)

    def read_ire(self):
        try:
            r = self.serial_port.readline().decode().strip()
            if r.isdigit():
                v = int(r)
                self.ire_label.config(text=f"IRE Level: {v}")
                self._highlight_ire(v)
        except: pass

    def _highlight_ire(self, v):
        key = min(self.ire_buttons.keys(), key=lambda k: abs(k-v))
        for k,btn in self.ire_buttons.items():
            btn.config(bg=self.on_bg if k==key else self.btn_bg)

    # — Pattern —

    def poll_pattern(self):
        if self.serial_port and self.serial_port.is_open:
            self.send_serial('J')
            self.master.after(100, self.read_pattern)
        self.master.after(20000, self.poll_pattern)

    def read_pattern(self):
        try:
            r = self.serial_port.readline().decode().strip()
            if r.isdigit():
                nm = self._pattern_map.get(int(r))
                if nm: self._highlight_pattern(nm)
        except: pass

    def select_pattern(self, n, nm):
        self.send_serial(f"{n}J")
        self._highlight_pattern(nm)

    def _highlight_pattern(self, chosen):
        for nm,btn in self.pattern_buttons.items():
            btn.config(bg=self.on_bg if nm==chosen else self.btn_bg)

    # — Resolution —

    def poll_resolution(self):
        if self.serial_port and self.serial_port.is_open:
            self.send_serial('=')
            self.master.after(100, self.read_resolution)
        self.master.after(20000, self.poll_resolution)

    def read_resolution(self):
        try:
            r = self.serial_port.readline().decode().strip()
            code = (re.search(r"(\d+\*\d+)", r) or [None, None])[1]
            if code:
                nm = self._res_map.get(code)
                if nm: self._highlight_resolution(nm)
        except: pass

    def select_resolution(self, cmd, nm):
        self.send_serial(cmd)
        self._highlight_resolution(nm)

    def _highlight_resolution(self, chosen):
        for nm,btn in self.res_buttons.items():
            btn.config(bg=self.on_bg if nm==chosen else self.btn_bg)

    # — Temperature —

    def poll_temperature(self):
        if self.serial_port and self.serial_port.is_open:
            self.send_serial('20S')
            self.master.after(100, self.read_temperature)
        self.master.after(60000, self.poll_temperature)

    def read_temperature(self):
        try:
            ttxt = self.serial_port.readline().decode().strip()
            m = re.search(r"([+-]?\d+\.?\d*)F", ttxt)
            if m:
                tf = float(m.group(1))
                self.temp_label.config(text=f"Temperature: {tf:.0f} °F")
        except: pass

    # — Color helpers —

    def select_color(self, cmd, nm):
        self.send_serial(cmd)
        self._highlight_color(nm)

    def _highlight_color(self, chosen):
        for nm, btn in self.color_buttons.items():
            if nm == chosen:
                btn.config(relief=tk.SUNKEN, bd=4)
            else:
                btn.config(relief=tk.RAISED, bd=2)

    # — HCFR integration —

    def toggle_hcfr(self):
        self.hcfr_active = bool(self.hcfr_var.get())
        if self.hcfr_active:
            self.master.after(int(POLL_INTERVAL*1000), self.hcfr_read)

    def hcfr_read(self):
        if not self.hcfr_active:
            return

        # grab combined UI/OCR text
        txt = ""
        for h in self.find_information_windows():
            ui = self.get_text_pywinauto(h) or ""
            oc = self.get_text_ocr(h) or ""
            txt += ui + "\n" + oc

        low = txt.lower()

        # color detection first
        for cue, cname in self.hcfr_color_map.items():
            if cue in low:
                cmd, _ = self.color_map[cname]
                self.select_color(cmd, cname)
                break
        else:
            # then IRE detection
            pct = self.parse_percentage(txt)
            if pct is not None:
                self.set_ire(pct)

        # always reschedule
        self.master.after(int(POLL_INTERVAL*1000), self.hcfr_read)

    @staticmethod
    def find_information_windows():
        wins = []
        def cb(h,l):
            if win32gui.IsWindowVisible(h) and "Information" in win32gui.GetWindowText(h):
                l.append(h)
        win32gui.EnumWindows(cb, wins)
        return wins

    @staticmethod
    def get_text_pywinauto(hwnd):
        try:
            dlg = Desktop(backend="uia").window(handle=hwnd)
            return "\n".join(c.window_text() for c in dlg.descendants() if c.window_text())
        except:
            return ""

    @staticmethod
    def get_text_ocr(hwnd):
        try:
            x1,y1,x2,y2 = win32gui.GetWindowRect(hwnd)
            img = ImageGrab.grab((x1,y1,x2,y2))
            return pytesseract.image_to_string(img).strip()
        except:
            return ""

    @staticmethod
    def parse_percentage(text):
        m = re.search(r"(\d{1,3})%\s*gray", text, re.IGNORECASE)
        if not m:
            return None
        v = int(m.group(1))
        if 0 <= v <= 100:
            return round(v/10)*10
        return None

if __name__ == "__main__":
    root = tk.Tk()
    CombinedController(root)
    root.mainloop()
