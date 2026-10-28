"""FreeConvert - portable, yerel Word <-> PDF dönüştürücü.
Tkinter arayüz. İnternet gerektirmez, veri cihazdan çıkmaz.
"""
from __future__ import annotations

import os
import json
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from converter.word_to_pdf import convert_word_to_pdf
from converter.pdf_to_word import convert_pdf_to_word
from converter.images import (
    convert_image_to_image,
    convert_images_to_pdf,
    convert_pdf_to_images,
)
from converter.ocr import convert_image_to_text, tesseract_available
from converter.pdf_tools import convert_pdfs_to_pdf, convert_pdf_to_text

APP_TITLE = "FreeConvert"

THEMES = {
    "light": {
        "bg": "#f0f0f0", "fg": "#000000",
        "field_bg": "#ffffff", "field_fg": "#000000",
        "select_bg": "#0078d7", "select_fg": "#ffffff",
        "hint_fg": "#555555", "privacy_fg": "#1a7f37",
        "button_bg": "#e1e1e1",
    },
    "dark": {
        "bg": "#2b2b2b", "fg": "#e8e8e8",
        "field_bg": "#3c3c3c", "field_fg": "#e8e8e8",
        "select_bg": "#0078d7", "select_fg": "#ffffff",
        "hint_fg": "#b0b0b0", "privacy_fg": "#4ecb71",
        "button_bg": "#3c3c3c",
    },
}


def _settings_file() -> str | None:
    try:
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        d = os.path.join(base, "FreeConvert")
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, "settings.json")
    except Exception:
        return None

S = {
    "tr": {
        "subtitle": "%100 Yerel",
        "privacy": "Tüm işlemler bu cihazda yapılır. İnternete veri gönderilmez.",
        "direction": "Dönüştürme",
        "direction_help": "İşlem türünü seçin; ayrıntılı yön sonraki adımda sorulur.",
        "doc_convert": "Belge Dönüştür\nWord ⇄ PDF ve PDF araçları",
        "visual_convert": "Görsel Dönüştür\nGörsel, PDF ve OCR",
        "choose_doc_title": "Belge dönüştürme yönü",
        "choose_visual_title": "Görsel / OCR dönüştürme yönü",
        "choose_direction_help": "Kullanılacak işlemi seçin:",
        "active_direction": "Seçili işlem: {d}",
        "word2pdf": "Word → PDF (.docx)",
        "pdf2word": "PDF → Word (.pdf → .docx)",
        "img2img": "Görsel → Görsel (PNG / JPG)",
        "img2pdf": "Görseller → PDF (birleştir)",
        "pdf2img": "PDF → Görseller (sayfalar)",
        "img2txt": "Görsel → Metin (OCR)",
        "pdfmerge": "PDF → PDF (birleştir)",
        "pdf2txt": "PDF → Metin (.pdf → .txt)",
        "img_opts": "Görsel seçenekleri",
        "target_fmt": "Hedef:",
        "quality": "JPG kalitesi:",
        "dpi_label": "Çözünürlük (DPI):",
        "mode_frame": "Dahili dönüşüm motoru",
        "mode_local": "Dahili offline motor — Word, LibreOffice veya başka uygulama gerekmez.",
        "mode_adaptif": "⚠ Adaptif (uyumluluk için)",
        "mode_sabit": "Sabit",
        "mode_sadakat": "Maksimum sadakat",
        "tip_adaptif": "Uygulama içindeki açık kaynak motor kullanılır.",
        "tip_sabit": "Uygulama içindeki açık kaynak motor kullanılır.",
        "tip_sadakat": "Uygulama içindeki açık kaynak motor kullanılır.",
        "hint_wp": "Eski genel ipucu.",
        "hint_ip": "Eski genel ipucu.",
        "hint_w2p": "İpucu: .docx dosyaları aynı adla PDF'e çevrilir.",
        "hint_p2w": "İpucu: PDF, Word belgesine aktarılır; taranmış sayfalarda yerel OCR denenir.",
        "hint_i2i": "İpucu: png/jpg/webp/bmp → png veya jpg. Saydamlık jpg'de beyaza basılır.",
        "hint_i2p": "İpucu: görseller listedeki sırayla tek PDF'te birleşir.",
        "hint_p2i": "İpucu: PDF sayfaları ayrı görseller olarak çıkar (ör. belge_p1.png).",
        "hint_i2t": "İpucu: görseldeki yazıyı metne çevirir (yerel OCR).",
        "hint_p2m": "İpucu: listedeki sırayla PDF'ler tek dosyada birleşir.",
        "hint_p2t": "İpucu: PDF'teki yazılar .txt olarak çıkar; taranmış sayfalarda OCR denenir.",
        "cancel": "İPTAL",
        "cancelling": "İptal ediliyor...",
        "done_cancelled": "İptal edildi: {ok}/{total} dosya dönüştürüldü.",
        "ocr_on": "● Yerel OCR hazır.",
        "ocr_off": "○ Yerel OCR motoru bulunamadı.",
        "add": "Dosya Ekle",
        "clear": "Listeyi Temizle",
        "remove_sel": "Seçiliyi Kaldır",
        "tip_remove": "Dosyayı diskten silmez; yalnızca listeden çıkarır.",
        "open_source": "Dosyayı Aç",
        "explore_source": "Gezginde Göster",
        "outdir_btn": "Çıktı Klasörü...",
        "source_noselect": "Önce listeden bir veya daha fazla dosya seçin.",
        "source_one": "Bu işlem için listeden tek bir dosya seçin.",
        "source_missing": "Dosya bulunamadı:\n{p}",
        "out_label": "Çıktı:",
        "files_frame": "Dosyalar (sürükle-bırak veya Dosya Ekle)",
        "skipped": "Uygun olmayan {n} dosya atlandı.",
        "ready": "Hazır.",
        "converting": "Dönüştürülüyor...",
        "convert": "DÖNÜŞTÜR",
        "warn_title": "Uyarı",
        "warn_nofile": "Önce dosya ekleyin.",
        "warn_bad_files": "Seçili işlemle uyumsuz veya eksik {n} dosya var. İlk sorun:\n{p}",
        "outdir_dialog": "Çıktı klasörü seç",
        "file_dialog": "Dosya seç",
        "types_word": "Word",
        "types_pdf": "PDF",
        "types_img": "Görseller",
        "types_all": "Tümü",
        "n_files": "{n} dosya listede.",
        "list_cleared": "Liste temizlendi.",
        "done_title": "Bitti",
        "done_msg": "{ok}/{total} dosya dönüştürüldü.",
        "done_status": "Bitti: {ok}/{total} başarılı.",
        "lang": "Dil",
        "theme_tip": "Tema değiştir (açık/koyu)",
        "hist_frame": "Dönüştürme geçmişi (açmak için çift tıkla)",
        "hist_open": "Aç",
        "hist_folder": "Klasörü Aç",
        "hist_clear": "Geçmişi Temizle",
        "hist_noselect": "Önce geçmişten bir dosya seç.",
        "hist_missing": "Dosya bulunamadı (taşınmış/silinmiş olabilir):\n{p}",
        "hist_cleared": "Geçmiş temizlendi.",
        "save_as_title": "Farklı kaydet — dosya zaten var",
        "extracted_msg": "görsel çıkarıldı.",
    },
    "en": {
        "subtitle": "100% Local",
        "privacy": "Everything runs on this device. No data is sent to the internet.",
        "direction": "Conversion",
        "direction_help": "Choose a category; you will select the exact operation next.",
        "doc_convert": "Convert Document\nWord ⇄ PDF and PDF tools",
        "visual_convert": "Convert Image\nImages, PDF and OCR",
        "choose_doc_title": "Document conversion direction",
        "choose_visual_title": "Image / OCR conversion direction",
        "choose_direction_help": "Choose the operation to use:",
        "active_direction": "Selected: {d}",
        "word2pdf": "Word → PDF (.docx)",
        "pdf2word": "PDF → Word (.pdf → .docx)",
        "img2img": "Image → Image (PNG / JPG)",
        "img2pdf": "Images → PDF (merge)",
        "pdf2img": "PDF → Images (pages)",
        "img2txt": "Image → Text (OCR)",
        "pdfmerge": "PDF → PDF (merge)",
        "pdf2txt": "PDF → Text (.pdf → .txt)",
        "img_opts": "Image options",
        "target_fmt": "Target:",
        "quality": "JPG quality:",
        "dpi_label": "Resolution (DPI):",
        "mode_frame": "Built-in conversion engine",
        "mode_local": "Built-in offline engine — no Word, LibreOffice or other app required.",
        "mode_adaptif": "⚠ Adaptive (compatibility)",
        "mode_sabit": "Fixed",
        "mode_sadakat": "Maximum fidelity",
        "tip_adaptif": "Uses the open-source engine bundled in the app.",
        "tip_sabit": "Uses the open-source engine bundled in the app.",
        "tip_sadakat": "Uses the open-source engine bundled in the app.",
        "hint_w2p": "Tip: .docx files are saved as PDF with the same name in the output folder.",
        "hint_p2w": "Tip: PDF is converted to Word; scanned pages use local OCR when available.",
        "hint_i2i": "Tip: png/jpg/webp/bmp → png or jpg. Transparency becomes white in jpg.",
        "hint_i2p": "Tip: images merge into one PDF in list order.",
        "hint_p2i": "Tip: each page becomes a separate image (e.g. doc_p1.png).",
        "hint_i2t": "Tip: reads text from images (local OCR).",
        "hint_p2m": "Tip: PDFs merge into one file in list order.",
        "hint_p2t": "Tip: extracts PDF text to .txt; scanned pages are OCR'd when available.",
        "cancel": "CANCEL",
        "cancelling": "Cancelling...",
        "done_cancelled": "Cancelled: {ok}/{total} file(s) converted.",
        "ocr_on": "● Local OCR ready.",
        "ocr_off": "○ Local OCR engine was not found.",
        "add": "Add Files",
        "clear": "Clear List",
        "remove_sel": "Remove Selected",
        "tip_remove": "Does not delete the file; only removes it from the list.",
        "open_source": "Open File",
        "explore_source": "Show in Explorer",
        "outdir_btn": "Output Folder...",
        "source_noselect": "Select one or more files from the list first.",
        "source_one": "Select exactly one file from the list for this action.",
        "source_missing": "File not found:\n{p}",
        "out_label": "Output:",
        "files_frame": "Files (drag & drop or Add Files)",
        "ready": "Ready.",
        "converting": "Converting...",
        "convert": "CONVERT",
        "warn_title": "Warning",
        "warn_nofile": "Please add files first.",
        "warn_bad_files": "{n} file(s) are missing or incompatible with the selected operation. First issue:\n{p}",
        "outdir_dialog": "Select output folder",
        "file_dialog": "Select files",
        "types_word": "Word",
        "types_pdf": "PDF",
        "types_img": "Images",
        "types_all": "All files",
        "n_files": "{n} file(s) in list.",
        "list_cleared": "List cleared.",
        "done_title": "Done",
        "done_msg": "{ok}/{total} file(s) converted.",
        "done_status": "Done: {ok}/{total} succeeded.",
        "lang": "Language",
        "theme_tip": "Toggle theme (light/dark)",
        "hist_frame": "Conversion history (double-click to open)",
        "hist_open": "Open",
        "hist_folder": "Open Folder",
        "hist_clear": "Clear History",
        "hist_noselect": "Select a file from history first.",
        "hist_missing": "File not found (may have been moved/deleted):\n{p}",
        "hist_cleared": "History cleared.",
        "save_as_title": "Save as — file already exists",
        "extracted_msg": "image(s) extracted.",
    },
}


def _set_app_user_model_id() -> None:
    """Görev çubuğunda Python yerine FreeConvert simgesinin gözükmesini sağlar."""
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "FreeConvert.App")
    except Exception:
        pass


def _apply_window_icon(window: tk.Tk) -> None:
    """Pencere başlığı ve görev çubuğu için logo; exe içindeki .ico kullanılır."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    ico = os.path.join(base, "assets", "icon", "freeconvert.ico")
    try:
        if os.name == "nt" and os.path.isfile(ico):
            window.iconbitmap(default=ico)
            window.after_idle(_set_native_icons, window, ico)
            return
        png = os.path.join(base, "assets", "icon", "freeconvert_1024.png")
        if os.path.isfile(png):
            window.iconphoto(True, tk.PhotoImage(file=png))
    except Exception:
        pass


def _set_native_icons(window: tk.Tk, ico: str) -> None:
    """Tk yalnızca 32 px ikon verir; görev çubuğu için 256 px kareyi doğrudan yükler."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p,
                                      ctypes.c_uint, ctypes.c_int,
                                      ctypes.c_int, ctypes.c_uint]
        user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                        ctypes.c_void_p, ctypes.c_void_p]
        image_icon, load_from_file = 1, 0x10
        wm_seticon, icon_small, icon_big = 0x80, 0, 1
        hwnd = user32.GetParent(window.winfo_id()) or window.winfo_id()
        big = user32.LoadImageW(None, ico, image_icon, 256, 256, load_from_file)
        small = user32.LoadImageW(None, ico, image_icon, 32, 32, load_from_file)
        if big:
            user32.SendMessageW(hwnd, wm_seticon, icon_big, big)
        if small:
            user32.SendMessageW(hwnd, wm_seticon, icon_small, small)
    except Exception:
        pass


def _ensure_start_menu_shortcut() -> None:
    """Windows Arama'da bulunsun diye Başlat menüsü kısayolunu yeniden yazar.

    Yalnızca paketlenmiş exe'de ve Windows'ta çalışır. Arka planda yapılır,
    açılışı geciktirmez.
    """
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return
    exe = os.path.abspath(sys.executable)
    link = os.path.join(appdata, "Microsoft", "Windows", "Start Menu",
                        "Programs", "FreeConvert.lnk")

    def _quote(value: str) -> str:
        return value.replace("'", "''")

    script = (
        f"$s = (New-Object -ComObject WScript.Shell).CreateShortcut('{_quote(link)}');"
        f"$s.TargetPath = '{_quote(exe)}';"
        f"$s.WorkingDirectory = '{_quote(os.path.dirname(exe))}';"
        f"$s.IconLocation = '{_quote(exe)},0';"
        "$s.Description = 'FreeConvert';"
        "$s.Save()"
    )

    def _write() -> None:
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive",
                 "-Command", script],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=60, check=False)
        except Exception:
            pass

    threading.Thread(target=_write, daemon=True).start()


class ToolTip:
    """Basit tooltip: imleç üzerine gelince açıklama gösterir."""

    def __init__(self, widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def update_text(self, text: str) -> None:
        self.text = text
        self._hide()

    def _show(self, _event=None) -> None:
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(self.tip, text=self.text, justify="left",
                       background="#ffffe0", relief="solid", borderwidth=1,
                       font=("Segoe UI", 9), wraplength=340, padx=6, pady=4)
        lbl.pack()

    def _hide(self, _event=None) -> None:
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


def _history_file() -> str | None:
    """Geçmiş JSON yolu. Yazılamazsa None (sadece bellek)."""
    try:
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        d = os.path.join(base, "FreeConvert")
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, "history.json")
    except Exception:
        return None


def _open_path(path: str) -> None:
    if hasattr(os, "startfile"):
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", path])


try:
    from tkinterdnd2 import TkinterDnD  # type: ignore
    _Base = TkinterDnD.Tk  # sürükle-bırak destekli
except Exception:
    _Base = tk.Tk


class App(_Base):  # type: ignore
    def __init__(self) -> None:
        super().__init__()
        _apply_window_icon(self)
        self.lang = tk.StringVar(value="tr")
        self.theme = tk.StringVar(value="light")
        self._load_settings()
        self._tooltips: dict[str, ToolTip] = {}
        self.files: list[str] = []
        self.history: list[dict] = []  # {src, dst, time}
        self.outdir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.mode = tk.StringVar(value="adaptif")
        self.direction = tk.StringVar(value="word2pdf")
        self.img_fmt = tk.StringVar(value="png")
        self.jpg_q = tk.IntVar(value=90)
        self.img_dpi = tk.IntVar(value=150)
        self._apply_settings()
        self._cancel_evt = threading.Event()
        self._converting = False
        self._q: queue.Queue = queue.Queue()
        self._load_history()

        # Kaydırılabilir içerik + sabit alt çubuk (DÖNÜŞTÜR hep görünür)
        self._container = ttk.Frame(self)
        self._container.pack(fill="both", expand=True)
        self._bottom = ttk.Frame(self._container, padding=(12, 0, 12, 12))
        self._bottom.pack(side="bottom", fill="x")
        scroll_host = ttk.Frame(self._container)
        scroll_host.pack(side="top", fill="both", expand=True)
        self._canvas = tk.Canvas(scroll_host, highlightthickness=0)
        self._vsb = ttk.Scrollbar(scroll_host, orient="vertical",
                                  command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vsb.set)
        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._scroll_inner = ttk.Frame(self._canvas, padding=12)
        self._win = self._canvas.create_window((0, 0),
                                               window=self._scroll_inner,
                                               anchor="nw")
        self._scroll_inner.bind(
            "<Configure>",
            lambda _e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfig(self._win, width=e.width))
        self._canvas.bind_all("<MouseWheel>", self._wheel, add="+")
        self._build_ui()
        self.after(150, self._poll_queue)

    def _wheel(self, event) -> None:
        try:
            w = event.widget
            if w is self._canvas or w is self._scroll_inner:
                self._canvas.yview_scroll(int(-event.delta / 120), "units")
                return
            if w.winfo_class() in ("Listbox", "Text", "TScale"):
                return  # kendi kaydırır
            self._canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:
            pass

    def t(self, key: str) -> str:
        return S[self.lang.get()].get(key, key)

    # ---- ayarlar (dil + tema + son kullanım kalıcılığı) ----
    _DIRECTIONS = ("word2pdf", "pdf2word", "img2img", "img2pdf",
                   "pdf2img", "img2txt", "pdfmerge", "pdf2txt")
    _DIRECTION_GROUPS = {
        "document": ("word2pdf", "pdf2word", "pdfmerge", "pdf2txt"),
        "visual": ("img2img", "img2pdf", "pdf2img", "img2txt"),
    }
    _FILE_EXTENSIONS = {
        "word2pdf": (".docx",),
        "pdf2word": (".pdf",),
        "pdfmerge": (".pdf",),
        "pdf2txt": (".pdf",),
        "img2img": (".png", ".jpg", ".jpeg", ".webp", ".bmp",
                    ".tif", ".tiff"),
        "img2pdf": (".png", ".jpg", ".jpeg", ".webp", ".bmp",
                    ".tif", ".tiff"),
        "img2txt": (".png", ".jpg", ".jpeg", ".webp", ".bmp",
                    ".tif", ".tiff"),
        "pdf2img": (".pdf",),
    }

    def _load_settings(self) -> None:
        self._settings_data: dict = {}
        p = _settings_file()
        if not p or not os.path.exists(p):
            return
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                self._settings_data = data
                if data.get("lang") in ("tr", "en"):
                    self.lang.set(data["lang"])
                if data.get("theme") in ("light", "dark"):
                    self.theme.set(data["theme"])
        except Exception:
            pass

    def _apply_settings(self) -> None:
        """Kayıtlı yön/mod/format/çıktı klasörünü geri yükle."""
        d = getattr(self, "_settings_data", {})
        if d.get("direction") in self._DIRECTIONS:
            self.direction.set(d["direction"])
        if d.get("mode") in ("adaptif", "sabit", "sadakat"):
            self.mode.set(d["mode"])
        if d.get("img_fmt") in ("png", "jpg"):
            self.img_fmt.set(d["img_fmt"])
        od = d.get("outdir")
        if isinstance(od, str) and od.strip() and os.path.isdir(od):
            self.outdir.set(od)

    def _save_settings(self) -> None:
        p = _settings_file()
        if not p:
            return
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump({"lang": self.lang.get(),
                           "theme": self.theme.get(),
                           "direction": self.direction.get(),
                           "mode": self.mode.get(),
                           "img_fmt": self.img_fmt.get(),
                           "outdir": self.outdir.get()}, f)
        except Exception:
            pass

    # ---- tema ----
    def _apply_theme(self) -> None:
        th = THEMES.get(self.theme.get(), THEMES["light"])
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=th["bg"])
        style.configure("TLabel", background=th["bg"],
                        foreground=th["fg"])
        style.configure("TLabelframe", background=th["bg"],
                        foreground=th["fg"])
        style.configure("TLabelframe.Label", background=th["bg"],
                        foreground=th["fg"])
        style.configure("TRadiobutton", background=th["bg"],
                        foreground=th["fg"])
        style.configure("TButton", background=th["button_bg"],
                        foreground=th["fg"])
        style.map("TButton",
                  background=[("active", th["select_bg"])])
        style.configure("TEntry", fieldbackground=th["field_bg"],
                        foreground=th["field_fg"])
        style.configure("TProgressbar", background=th["select_bg"])
        try:
            self.configure(background=th["bg"])
        except Exception:
            pass
        try:
            self._canvas.configure(background=th["bg"])
        except Exception:
            pass
        for lb in (getattr(self, "listbox", None),
                   getattr(self, "histbox", None)):
            if lb is not None:
                try:
                    lb.configure(background=th["field_bg"],
                                 foreground=th["field_fg"],
                                 selectbackground=th["select_bg"],
                                 selectforeground=th["select_fg"])
                except Exception:
                    pass
        if hasattr(self, "privacy_label"):
            self.privacy_label.configure(background=th["bg"],
                                         foreground=th["privacy_fg"])
        if hasattr(self, "hint"):
            self.hint.configure(background=th["bg"],
                                foreground=th["hint_fg"])
        if hasattr(self, "direction_summary"):
            self.direction_summary.configure(background=th["bg"],
                                            foreground=th["select_bg"])
        if hasattr(self, "theme_btn"):
            self.theme_btn.configure(
                text="☀" if self.theme.get() == "dark" else "☾")

    def toggle_theme(self) -> None:
        self.theme.set("dark" if self.theme.get() == "light" else "light")
        self._save_settings()
        self._apply_theme()

    # ---- UI ----
    def _build_ui(self) -> None:
        self._tooltips.clear()
        for w in self._scroll_inner.winfo_children():
            w.destroy()
        for w in self._bottom.winfo_children():
            w.destroy()
        top = self._scroll_inner
        bottom = self._bottom
        self._input_buttons: list[ttk.Button] = []

        self.title(f"{APP_TITLE} - {self.t('subtitle')}")
        if not getattr(self, "_sized", False):
            self.geometry("680x780")
            self._sized = True
        self.minsize(560, 480)

        header = ttk.Frame(top)
        header.pack(fill="x")
        ttk.Label(header, text=APP_TITLE,
                  font=("Segoe UI", 16, "bold")).pack(side="left")
        lang_frame = ttk.Frame(header)
        lang_frame.pack(side="right")
        ttk.Label(lang_frame, text=self.t("lang") + ":").pack(side="left", padx=(0, 4))
        ttk.Radiobutton(lang_frame, text="TR", value="tr",
                        variable=self.lang,
                        command=self._rebuild).pack(side="left")
        ttk.Radiobutton(lang_frame, text="EN", value="en",
                        variable=self.lang,
                        command=self._rebuild).pack(side="left", padx=(4, 0))
        self.theme_btn = ttk.Button(lang_frame, text="☾", width=3,
                                    command=self.toggle_theme)
        self.theme_btn.pack(side="left", padx=(8, 0))
        ToolTip(self.theme_btn, self.t("theme_tip"))

        self.privacy_label = ttk.Label(top, text=self.t("privacy"),
                                       foreground="#1a7f37")
        self.privacy_label.pack(anchor="w", pady=(0, 8))

        self.direction_frame = ttk.LabelFrame(
            top, text=self.t("direction"), padding=8)
        self.direction_frame.pack(fill="x", pady=4)
        ttk.Label(self.direction_frame, text=self.t("direction_help")).pack(
            anchor="w", pady=(0, 6))
        category_row = ttk.Frame(self.direction_frame)
        category_row.pack(fill="x")
        doc_direction_btn = ttk.Button(
            category_row, text=self.t("doc_convert"),
            command=lambda: self._choose_direction("document"))
        doc_direction_btn.pack(side="left", fill="both", expand=True,
                               padx=(0, 4))
        visual_direction_btn = ttk.Button(
            category_row, text=self.t("visual_convert"),
            command=lambda: self._choose_direction("visual"))
        visual_direction_btn.pack(side="left", fill="both", expand=True,
                                  padx=(4, 0))
        self._input_buttons.extend((doc_direction_btn, visual_direction_btn))
        self.direction_summary = ttk.Label(
            top, text="", foreground="#555", wraplength=610)
        self.direction_summary.pack(anchor="w", pady=(2, 4))

        self.mode_frame = ttk.LabelFrame(top, text=self.t("mode_frame"), padding=8)
        self.mode_frame.pack(fill="x", pady=4)
        mode_label = ttk.Label(self.mode_frame, text=self.t("mode_local"))
        mode_label.pack(anchor="w")
        self._tooltips["mode_local"] = ToolTip(mode_label, self.t("mode_local"))

        self.img_frame = ttk.LabelFrame(top, text=self.t("img_opts"), padding=8)
        self.img_frame.pack(fill="x", pady=4)
        fmt_row = ttk.Frame(self.img_frame)
        fmt_row.pack(fill="x")
        ttk.Label(fmt_row, text=self.t("target_fmt")).pack(side="left")
        ttk.Radiobutton(fmt_row, text="PNG", value="png",
                        variable=self.img_fmt,
                        command=self._refresh_q).pack(side="left", padx=4)
        ttk.Radiobutton(fmt_row, text="JPG", value="jpg",
                        variable=self.img_fmt,
                        command=self._refresh_q).pack(side="left", padx=4)
        self.q_row = ttk.Frame(self.img_frame)
        self.q_row.pack(fill="x", pady=(4, 0))
        self.q_label = ttk.Label(self.q_row, text="")
        self.q_label.pack(side="left")
        self.q_scale = ttk.Scale(self.q_row, from_=10, to=100,
                                 variable=self.jpg_q, orient="horizontal",
                                 length=140,
                                 command=lambda _v: self._refresh_q())
        self.q_scale.pack(side="left", padx=6)
        self.dpi_row = ttk.Frame(self.img_frame)
        self.dpi_row.pack(fill="x", pady=(4, 0))
        self.dpi_label = ttk.Label(self.dpi_row, text="")
        self.dpi_label.pack(side="left")
        self.dpi_scale = ttk.Scale(self.dpi_row, from_=72, to=300,
                                   variable=self.img_dpi, orient="horizontal",
                                   length=140,
                                   command=lambda _v: self._refresh_q())
        self.dpi_scale.pack(side="left", padx=6)

        self.hint = ttk.Label(top, text="", foreground="#555")
        self.hint.pack(anchor="w", pady=(2, 4))

        out_row = ttk.Frame(top)
        out_row.pack(fill="x", pady=(2, 4))
        ttk.Label(out_row, text=self.t("out_label")).pack(side="left")
        ttk.Entry(out_row, textvariable=self.outdir,
                  width=60).pack(side="left", padx=6, fill="x", expand=True)

        list_frame = ttk.LabelFrame(top, text=self.t("files_frame"), padding=6)
        list_frame.pack(fill="both", expand=True, pady=6)

        file_btns = ttk.Frame(list_frame)
        file_btns.pack(fill="x", pady=(0, 5))
        file_btn_row1 = ttk.Frame(file_btns)
        file_btn_row1.pack(fill="x", pady=(0, 3))
        file_btn_row2 = ttk.Frame(file_btns)
        file_btn_row2.pack(fill="x")

        add_btn = ttk.Button(file_btn_row1, text=self.t("add"),
                             command=self.add_files)
        add_btn.pack(side="left", fill="x", expand=True, padx=(0, 3))
        remove_btn = ttk.Button(file_btn_row1, text=self.t("remove_sel"),
                                command=self.remove_selected)
        remove_btn.pack(side="left", fill="x", expand=True, padx=3)
        clear_btn = ttk.Button(file_btn_row1, text=self.t("clear"),
                               command=self.clear_files)
        clear_btn.pack(side="left", fill="x", expand=True, padx=(3, 0))
        self._tooltips["remove"] = ToolTip(remove_btn, self.t("tip_remove"))

        open_btn = ttk.Button(file_btn_row2, text=self.t("open_source"),
                              command=self.open_selected_files)
        open_btn.pack(side="left", fill="x", expand=True, padx=(0, 3))
        explore_btn = ttk.Button(file_btn_row2, text=self.t("explore_source"),
                                 command=self.explore_selected_files)
        explore_btn.pack(side="left", fill="x", expand=True, padx=3)
        outdir_btn = ttk.Button(file_btn_row2, text=self.t("outdir_btn"),
                                command=self.pick_outdir)
        outdir_btn.pack(side="left", fill="x", expand=True, padx=(3, 0))
        self._input_buttons.extend(
            (add_btn, remove_btn, clear_btn, open_btn, explore_btn, outdir_btn))

        file_list_frame = ttk.Frame(list_frame)
        file_list_frame.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(file_list_frame, height=7,
                                 selectmode="extended")
        self.listbox.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(file_list_frame, orient="vertical",
                           command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.bind("<Delete>", lambda _e: self.remove_selected())
        self.listbox.bind("<Double-Button-1>", self.open_selected_files)
        for p in self.files:
            self.listbox.insert("end", p)

        # Sürükle-bırak (tkinterdnd2 varsa)
        try:
            from tkinterdnd2 import DND_FILES  # type: ignore
            assert hasattr(self, "drop_target_register")
            self.drop_target_register(DND_FILES)  # type: ignore
            self.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore
        except Exception:
            pass

        hist_frame = ttk.LabelFrame(top, text=self.t("hist_frame"), padding=6)
        hist_frame.pack(fill="x", pady=6)
        self.histbox = tk.Listbox(hist_frame, height=5, selectmode="extended")
        self.histbox.pack(fill="x", expand=True, side="left")
        hsb = ttk.Scrollbar(hist_frame, orient="vertical",
                            command=self.histbox.yview)
        hsb.pack(side="right", fill="y")
        self.histbox.config(yscrollcommand=hsb.set)
        self.histbox.bind("<Double-Button-1>", self.open_selected_history)
        self._refresh_history_listbox()

        hist_btns = ttk.Frame(top)
        hist_btns.pack(fill="x", pady=(0, 4))
        ttk.Button(hist_btns, text=self.t("hist_open"),
                   command=self.open_selected_history).pack(side="left", padx=2)
        ttk.Button(hist_btns, text=self.t("hist_folder"),
                   command=self.open_output_folder).pack(side="left", padx=2)
        ttk.Button(hist_btns, text=self.t("hist_clear"),
                   command=self.clear_history).pack(side="left", padx=2)

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x", pady=4)
        self.status = ttk.Label(bottom, text=self.t("ready"))
        self.status.pack(anchor="w")

        self.convert_btn = ttk.Button(bottom, text=self.t("convert"),
                                      command=self.start_convert)
        self.convert_btn.pack(fill="x", pady=8)
        self._set_converting(bool(getattr(self, "_converting", False)))

        self._ocr_ok = tesseract_available()
        self._refresh_hint()
        self._refresh_q()
        self._refresh_options()
        self._refresh_direction_summary()
        self._apply_theme()
        self._canvas.configure(
            scrollregion=self._canvas.bbox("all"))

    def _on_direction(self) -> None:
        self._save_settings()
        self._refresh_hint()
        self._refresh_options()
        self._refresh_direction_summary()

    def _refresh_direction_summary(self) -> None:
        if not hasattr(self, "direction_summary"):
            return
        label = self.t(self.direction.get())
        self.direction_summary.config(
            text=self.t("active_direction").format(d=label))

    def _choose_direction(self, group: str) -> None:
        """İki genel kategori butonundan sonra gerçek yönü sor."""
        keys = self._DIRECTION_GROUPS.get(group, ())
        if not keys:
            return

        dialog = tk.Toplevel(self)
        dialog.title(self.t("choose_doc_title" if group == "document"
                            else "choose_visual_title"))
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(background=THEMES.get(
            self.theme.get(), THEMES["light"])["bg"])

        body = ttk.Frame(dialog, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=self.t("choose_direction_help")).pack(
            anchor="w", pady=(0, 8))
        for key in keys:
            prefix = "✓ " if self.direction.get() == key else ""
            ttk.Button(
                body, text=prefix + self.t(key), width=42,
                command=lambda value=key: self._pick_modal_direction(
                    value, dialog)
            ).pack(fill="x", pady=3)
        ttk.Button(body, text=self.t("cancel"), width=42,
                   command=dialog.destroy).pack(fill="x", pady=(8, 0))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())

        dialog.update_idletasks()
        width = max(420, dialog.winfo_reqwidth())
        height = dialog.winfo_reqheight()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 3)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.grab_set()
        self.wait_window(dialog)
        try:
            self.focus_set()
        except Exception:
            pass

    def _pick_modal_direction(self, value: str, dialog: tk.Toplevel) -> None:
        if value not in self._DIRECTIONS:
            dialog.destroy()
            return
        self.direction.set(value)
        dialog.destroy()
        self._on_direction()

    def _refresh_q(self) -> None:
        if hasattr(self, "q_label"):
            self.q_label.config(
                text=f"{self.t('quality')} {int(self.jpg_q.get())}")
        if hasattr(self, "dpi_label"):
            self.dpi_label.config(
                text=f"{self.t('dpi_label')} {int(self.img_dpi.get())}")
        if hasattr(self, "img_frame"):
            self._refresh_options()

    def _refresh_options(self) -> None:
        if not hasattr(self, "img_frame"):
            return
        d = self.direction.get()
        # Dahili motor bilgisi yalnızca belge yönlerinde gösterilir.
        # Görsel/OCR yönlerinde ilgisiz kalite seçeneği gizlenir.
        if hasattr(self, "mode_frame"):
            if d in ("word2pdf", "pdf2word"):
                self.mode_frame.pack(fill="x", pady=4, before=self.hint)
            else:
                self.mode_frame.pack_forget()
        if d in ("img2img", "pdf2img"):
            self.img_frame.pack(fill="x", pady=4, before=self.hint)
            if d == "img2img":
                self.dpi_row.pack_forget()
            else:
                self.dpi_row.pack(fill="x", pady=(4, 0))
            # JPG kalitesi yalnızca JPG hedefi seçiliyse anlamlıdır.
            if self.img_fmt.get() == "jpg":
                self.q_row.pack(fill="x", pady=(4, 0))
            else:
                self.q_row.pack_forget()
        else:
            self.img_frame.pack_forget()

    def _rebuild(self) -> None:
        self._save_settings()
        self._build_ui()
        self._set_status(self.t("ready"))

    def _refresh_hint(self) -> None:
        d = self.direction.get()
        if d == "word2pdf":
            base = self.t("hint_w2p")
        elif d == "pdf2word":
            base = self.t("hint_p2w")
        elif d == "img2img":
            base = self.t("hint_i2i")
        elif d == "img2pdf":
            base = self.t("hint_i2p")
        elif d == "img2txt":
            base = self.t("hint_i2t")
        elif d == "pdfmerge":
            base = self.t("hint_p2m")
        elif d == "pdf2txt":
            base = self.t("hint_p2t")
        else:
            base = self.t("hint_p2i")
        try:
            ocr_ok = bool(self._ocr_ok)
        except AttributeError:
            ocr_ok = False
        text = base
        # OCR durumu yalnızca OCR'ın gerçekten kullanılabileceği işlemlerde
        # gösterilir; normal Word/görsel dönüşümlerinde ek gürültü yaratmaz.
        if d in ("img2txt", "pdf2txt", "pdf2word"):
            text += " " + self.t("ocr_on" if ocr_ok else "ocr_off")
        self.hint.config(text=text)

    # ---- dosya işlemleri ----
    def _is_compatible_file(self, path: str) -> bool:
        allowed = self._FILE_EXTENSIONS.get(self.direction.get(), ())
        return (os.path.isfile(path)
                and os.path.splitext(path)[1].lower() in allowed)

    def _add_source_paths(self, paths) -> tuple[int, int]:
        """Uygun ve yeni dosyaları ekler. Dönüş: (eklenen, atlanan)."""
        added = skipped = 0
        for path in paths:
            if not path or not self._is_compatible_file(path):
                skipped += 1
                continue
            if path not in self.files:
                self.files.append(path)
                self.listbox.insert("end", path)
                added += 1
        return added, skipped

    def _file_count_status(self, skipped: int = 0) -> str:
        text = self.t("n_files").format(n=len(self.files))
        if skipped:
            text += " " + self.t("skipped").format(n=skipped)
        return text

    def add_files(self) -> None:
        d = self.direction.get()
        if d == "word2pdf":
            types = [(self.t("types_word"), ("*.docx",)),
                     (self.t("types_all"), "*.*")]
        elif d in ("img2img", "img2pdf", "img2txt"):
            types = [(self.t("types_img"),
                      ("*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp",
                       "*.tif", "*.tiff")),
                     (self.t("types_all"), "*.*")]
        else:
            types = [(self.t("types_pdf"), ("*.pdf",)),
                     (self.t("types_all"), "*.*")]
        paths = filedialog.askopenfilenames(title=self.t("file_dialog"),
                                            filetypes=types)
        _added, skipped = self._add_source_paths(paths)
        self._set_status(self._file_count_status(skipped))

    def _on_drop(self, event) -> None:  # tkinterdnd2 event
        raw = str(getattr(event, "data", ""))
        # {C:\a b.pdf} {C:\c.pdf} formatını çöz
        parts, cur, in_brace = [], "", False
        for ch in raw:
            if ch == "{":
                in_brace = True
                cur = ""
            elif ch == "}":
                in_brace = False
                parts.append(cur)
                cur = ""
            elif ch == " " and not in_brace:
                if cur.strip():
                    parts.append(cur.strip())
                cur = ""
            else:
                cur += ch
        if cur.strip():
            parts.append(cur.strip())
        _added, skipped = self._add_source_paths(parts)
        self._set_status(self._file_count_status(skipped))

    def clear_files(self) -> None:
        self.files.clear()
        self.listbox.delete(0, "end")
        self._set_status(self.t("list_cleared"))

    def remove_selected(self) -> None:
        """Listeden seçili satırları kaldır (Delete tuşu da çalışır)."""
        if not hasattr(self, "listbox"):
            return
        sel = sorted(self.listbox.curselection(), reverse=True)
        for i in sel:
            self.listbox.delete(i)
            if 0 <= i < len(self.files):
                del self.files[i]
        self._set_status(self.t("n_files").format(n=len(self.files)))

    def _selected_source_paths(self, require_one: bool = False) -> list[str]:
        if not hasattr(self, "listbox"):
            return []
        selected = sorted(self.listbox.curselection())
        if not selected:
            messagebox.showwarning(self.t("warn_title"),
                                   self.t("source_noselect"))
            return []
        if require_one and len(selected) != 1:
            messagebox.showwarning(self.t("warn_title"),
                                   self.t("source_one"))
            return []
        paths: list[str] = []
        for index in selected:
            if 0 <= index < len(self.files):
                path = self.files[index]
                if os.path.isfile(path):
                    paths.append(path)
                else:
                    messagebox.showwarning(
                        self.t("warn_title"),
                        self.t("source_missing").format(p=path))
        return paths

    def open_selected_files(self, _event=None) -> None:
        """Listede seçili tek kaynak dosyayı varsayılan uygulamada aç."""
        paths = self._selected_source_paths(require_one=True)
        if not paths:
            return
        try:
            _open_path(paths[0])
        except Exception as ex:
            messagebox.showwarning(self.t("warn_title"), str(ex))

    def explore_selected_files(self) -> None:
        """Kaynak dosyayı Windows Gezgini'nde seçili gösterir."""
        paths = self._selected_source_paths()
        if not paths:
            return
        try:
            if len(paths) == 1 and os.name == "nt":
                subprocess.Popen(
                    ["explorer", f"/select,{os.path.normpath(paths[0])}"])
            else:
                folders = list(dict.fromkeys(
                    os.path.dirname(p) for p in paths if os.path.dirname(p)))
                for folder in folders:
                    _open_path(folder)
        except Exception as ex:
            messagebox.showwarning(self.t("warn_title"), str(ex))

    def pick_outdir(self) -> None:
        d = filedialog.askdirectory(title=self.t("outdir_dialog"))
        if d:
            self.outdir.set(d)
            self._save_settings()

    # ---- geçmiş ----
    def _load_history(self) -> None:
        self.history = []
        p = _history_file()
        if not p or not os.path.exists(p):
            return
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                self.history = [e for e in data
                                if isinstance(e, dict) and e.get("dst")]
        except Exception:
            self.history = []

    def _save_history(self) -> None:
        p = _history_file()
        if not p:
            return
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(self.history[-100:], f,
                          ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _add_history(self, src: str, dst: str) -> None:
        self.history.append({"src": src, "dst": dst,
                             "time": datetime.now().strftime("%d.%m %H:%M")})
        self.history = self.history[-100:]
        self._save_history()
        self._refresh_history_listbox()

    def _refresh_history_listbox(self) -> None:
        if not hasattr(self, "histbox"):
            return
        self.histbox.delete(0, "end")
        for e in self.history:
            ok_mark = "✓ " if os.path.exists(e.get("dst", "")) else "✗ "
            self.histbox.insert(
                "end",
                f"{ok_mark}{e.get('time', '')}  "
                f"{os.path.basename(e.get('dst', ''))}")

    def open_selected_history(self, _event=None) -> None:
        sel = self.histbox.curselection() if hasattr(self, "histbox") else ()
        if not sel:
            messagebox.showwarning(self.t("warn_title"),
                                   self.t("hist_noselect"))
            return
        for i in sel:
            try:
                dst = self.history[i]["dst"]
            except IndexError:
                continue
            if not os.path.exists(dst):
                messagebox.showwarning(
                    self.t("warn_title"),
                    self.t("hist_missing").format(p=dst))
                continue
            try:
                _open_path(dst)
            except Exception as ex:
                messagebox.showwarning(self.t("warn_title"), str(ex))

    def open_output_folder(self) -> None:
        d = self.outdir.get().strip() or os.getcwd()
        if os.path.isdir(d):
            try:
                _open_path(d)
                return
            except Exception as ex:
                messagebox.showwarning(self.t("warn_title"), str(ex))
                return
        messagebox.showwarning(self.t("warn_title"),
                               self.t("hist_missing").format(p=d))

    def clear_history(self) -> None:
        self.history.clear()
        self._save_history()
        self._refresh_history_listbox()
        self._set_status(self.t("hist_cleared"))

    def _set_status(self, msg: str) -> None:
        self.status.config(text=msg)

    # ---- çıktı adı çakışması ----
    def _resolve_collision(self, dst: str) -> str | None:
        """Çıktı varsa Farklı-Kaydet sor; vazgeçilirse None."""
        if not os.path.exists(dst):
            return dst
        new = filedialog.asksaveasfilename(
            title=self.t("save_as_title"),
            initialdir=os.path.dirname(dst) or os.getcwd(),
            initialfile=os.path.basename(dst),
            defaultextension=os.path.splitext(dst)[1],
            filetypes=[(self.t("types_all"), "*.*")])
        return new or None

    def _auto_name(self, dst: str) -> str:
        """Soru sormadan sıradaki boş adı bul (çok çıktılı işler için)."""
        if not os.path.exists(dst):
            return dst
        root, ext = os.path.splitext(dst)
        i = 2
        while os.path.exists(f"{root}_{i}{ext}"):
            i += 1
        return f"{root}_{i}{ext}"

    def _plan_pdf_images(self, src: str, outdir: str,
                         fmt: str) -> list[str] | None:
        """PDF sayfaları için çıktı adlarını önceden planla.

        Açılamazsa None (worker hatayı raporlar, adları kendi üretir).
        """
        try:
            import pypdfium2 as pdfium
            doc = pdfium.PdfDocument(src)
            n = len(doc)
            doc.close()
        except Exception:
            return None
        ext = ".png" if fmt == "png" else ".jpg"
        base = os.path.splitext(os.path.basename(src))[0]
        return [self._auto_name(os.path.join(outdir, f"{base}_p{i}{ext}"))
                for i in range(1, n + 1)]

    # ---- dönüştürme ----
    def start_convert(self) -> None:
        if not self.files:
            messagebox.showwarning(self.t("warn_title"),
                                   self.t("warn_nofile"))
            return
        outdir = self.outdir.get().strip() or os.getcwd()
        os.makedirs(outdir, exist_ok=True)
        direction = self.direction.get()
        mode = self.mode.get()
        lang = self.lang.get()
        img_fmt = self.img_fmt.get()
        jpg_q = int(self.jpg_q.get())
        img_dpi = int(self.img_dpi.get())
        files = list(self.files)
        bad_files = [p for p in files if not self._is_compatible_file(p)]
        if bad_files:
            messagebox.showwarning(
                self.t("warn_title"),
                self.t("warn_bad_files").format(
                    n=len(bad_files), p=bad_files[0]))
            return
        self._save_settings()  # yön/mod/format son haliyle kaydet

        tasks: list[tuple] = []  # ("one", src, dst) | ("merge", srcs, dst)
        #                                | ("split", src, dsts|None)
        if direction in ("img2pdf", "pdfmerge"):
            base = os.path.splitext(os.path.basename(files[0]))[0]
            if len(files) > 1:
                suffix = "_combined" if lang == "en" else "_birlesik"
                dst = os.path.join(outdir, base + suffix + ".pdf")
            else:
                dst = os.path.join(outdir, base + ".pdf")
            dst = self._resolve_collision(dst)
            if dst is None:
                return
            tasks.append(("merge", files, dst))
        else:
            for src in files:
                base = os.path.splitext(os.path.basename(src))[0]
                if direction == "word2pdf":
                    dst = self._resolve_collision(
                        os.path.join(outdir, base + ".pdf"))
                    if dst is None:
                        continue
                    tasks.append(("one", src, dst))
                elif direction == "pdf2word":
                    dst = self._resolve_collision(
                        os.path.join(outdir, base + ".docx"))
                    if dst is None:
                        continue
                    tasks.append(("one", src, dst))
                elif direction == "img2img":
                    ext = ".png" if img_fmt == "png" else ".jpg"
                    dst = self._resolve_collision(
                        os.path.join(outdir, base + ext))
                    if dst is None:
                        continue
                    tasks.append(("one", src, dst))
                elif direction == "img2txt":
                    dst = self._resolve_collision(
                        os.path.join(outdir, base + ".txt"))
                    if dst is None:
                        continue
                    tasks.append(("one", src, dst))
                elif direction == "pdf2txt":
                    dst = self._resolve_collision(
                        os.path.join(outdir, base + ".txt"))
                    if dst is None:
                        continue
                    tasks.append(("one", src, dst))
                else:  # pdf2img: çok çıktı -> otomatik ad, sorusuz
                    tasks.append(("split", src,
                                  self._plan_pdf_images(src, outdir, img_fmt)))
        if not tasks:
            return
        self.progress.config(maximum=len(tasks), value=0)
        self._set_status(self.t("converting"))
        self._cancel_evt.clear()
        self._set_converting(True)
        threading.Thread(target=self._worker,
                         args=(tasks, outdir, direction, mode, lang,
                               img_fmt, jpg_q, img_dpi),
                         daemon=True).start()

    def _worker(self, tasks: list[tuple], outdir: str, direction: str,
                mode: str, lang: str, img_fmt: str, jpg_q: int,
                img_dpi: int) -> None:
        tag_err = "ERROR" if lang == "en" else "HATA"
        ok_count = 0
        cancelled = False
        for i, task in enumerate(tasks, 1):
            if self._cancel_evt.is_set():
                cancelled = True
                break
            kind = task[0]
            if kind == "merge":
                _, srcs, dst = task
                if direction == "pdfmerge":
                    ok, info = convert_pdfs_to_pdf(srcs, dst)
                else:
                    ok, info = convert_images_to_pdf(srcs, dst)
                outs = [dst] if ok else []
                label = os.path.basename(dst)
            elif kind == "split":
                _, src, dsts = task
                ok, result = convert_pdf_to_images(
                    src, outdir, fmt=img_fmt, dpi=img_dpi, quality=jpg_q,
                    out_names=dsts)
                if ok:
                    outs = result
                    info = f"{len(outs)} {S[lang]['extracted_msg']}"
                else:
                    outs = []
                    info = result
                label = os.path.basename(src)
            else:  # one
                _, src, dst = task
                if direction == "word2pdf":
                    ok, info = convert_word_to_pdf(src, dst, mode=mode)
                elif direction == "pdf2word":
                    ok, info = convert_pdf_to_word(src, dst, mode=mode)
                elif direction == "img2img":
                    ok, info = convert_image_to_image(src, dst,
                                                      quality=jpg_q)
                elif direction == "pdf2txt":
                    ok, info = convert_pdf_to_text(src, dst)
                else:  # img2txt
                    ok, info = convert_image_to_text(src, dst)
                outs = [dst] if ok else []
                label = os.path.basename(src)
            ok_count += 1 if ok else 0
            tag = "OK" if ok else tag_err
            self._q.put(("progress", i, len(tasks),
                         f"[{tag}] {label} → {info}"))
            for o in outs:
                self._q.put(("history", src if kind != "merge" else ";".join(srcs), o))
        self._q.put(("done", ok_count, len(tasks), lang, cancelled))

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self._q.get_nowait()
                if msg[0] == "progress":
                    _, i, total, text = msg
                    self.progress.config(value=i)
                    self._set_status(f"{i}/{total} {text}")
                elif msg[0] == "history":
                    _, src, dst = msg
                    self._add_history(src, dst)
                elif msg[0] == "done":
                    _, ok_count, total, lang, cancelled = msg
                    s = S.get(lang, S["tr"])
                    self._set_converting(False)
                    self._cancel_evt.clear()
                    if cancelled:
                        txt = s["done_cancelled"].format(
                            ok=ok_count, total=total)
                        self._set_status(txt)
                        messagebox.showinfo(s["done_title"], txt)
                    else:
                        self._set_status(s["done_status"].format(
                            ok=ok_count, total=total))
                        messagebox.showinfo(
                            s["done_title"],
                            s["done_msg"].format(ok=ok_count, total=total))
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    # ---- dönüşüm sırasında iptal ----
    def _set_converting(self, flag: bool) -> None:
        """DÖNÜŞTÜR düğmesini iptal moduna al / geri getir."""
        self._converting = flag
        if not hasattr(self, "convert_btn"):
            return
        if flag:
            self.convert_btn.config(text=self.t("cancel"),
                                    command=self._request_cancel)
        else:
            self.convert_btn.config(text=self.t("convert"),
                                    command=self.start_convert)
        for button in getattr(self, "_input_buttons", []):
            try:
                button.state(["disabled"] if flag else ["!disabled"])
            except Exception:
                pass

    def _request_cancel(self) -> None:
        self._cancel_evt.set()
        try:
            self._set_status(self.t("cancelling"))
        except Exception:
            pass


def _run_self_test(args: list[str]) -> int:
    """Portable frozen build'ı konsol açmadan doğrulamak için test girişi."""
    if len(args) != 4:
        return 2
    kind, input_path, output_path, result_path = args
    operations = {
        "word2pdf": convert_word_to_pdf,
        "pdf2word": convert_pdf_to_word,
        "img2txt": convert_image_to_text,
        "pdf2txt": convert_pdf_to_text,
    }
    operation = operations.get(kind)
    payload: dict[str, object] = {"kind": kind, "ok": False}
    try:
        if operation is None:
            raise ValueError("Unknown self-test operation")
        ok, info = operation(input_path, output_path)
        payload.update({"ok": bool(ok), "info": str(info),
                        "output_exists": os.path.exists(output_path)})
    except Exception as exc:
        payload["error"] = str(exc)
    try:
        os.makedirs(os.path.dirname(os.path.abspath(result_path)),
                    exist_ok=True)
        with open(result_path, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False)
    except Exception:
        return 3
    return 0 if payload.get("ok") and payload.get("output_exists") else 1


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--self-test":
        raise SystemExit(_run_self_test(sys.argv[2:]))
    _set_app_user_model_id()
    _ensure_start_menu_shortcut()
    app = App()
    app.mainloop()
