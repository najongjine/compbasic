import json
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import parse_qs, urlsplit


APP_NAME = "유튜브 다운로드"
APP_DATA = Path(os.environ["LOCALAPPDATA"]) / "YouTubeDL"
SETTINGS = APP_DATA / "settings.json"
DOWNLOADER = APP_DATA / "venv" / "Scripts" / "python.exe"
DEFAULT_FOLDER = Path.home() / "Downloads" / "YouTube"
QUALITIES = ("최고 화질", "2160p (4K)", "1440p", "1080p", "720p", "480p", "360p")
DEFAULT_QUALITY = "1080p"
VIDEO_ID = re.compile(r"[A-Za-z0-9_-]{11}")
YOUTUBE_HOSTS = {"youtube.com", "www.youtube.com", "m.youtube.com", "music.youtube.com"}


def video_url(value):
    value = value.strip()
    if "://" not in value:
        value = f"https://{value}"
    url = urlsplit(value)
    parts = url.path.strip("/").split("/")
    video_id = ""
    if url.scheme in {"http", "https"} and not url.username and not url.password:
        if url.hostname in {"youtu.be", "www.youtu.be"} and len(parts) == 1:
            video_id = parts[0]
        elif url.hostname in YOUTUBE_HOSTS:
            if parts == ["watch"]:
                video_id = parse_qs(url.query).get("v", [""])[0]
            elif len(parts) == 2 and parts[0] in {"shorts", "live", "embed", "v"}:
                video_id = parts[1]
    if not VIDEO_ID.fullmatch(video_id):
        raise ValueError("유튜브 영상 주소를 붙여넣어 주세요.")
    return f"https://www.youtube.com/watch?v={video_id}"


def load_settings():
    try:
        settings = json.loads(SETTINGS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return settings if isinstance(settings, dict) else {}


def save_settings(folder, quality):
    APP_DATA.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=APP_DATA, delete=False
    ) as temporary:
        path = Path(temporary.name)
        try:
            json.dump({"folder": str(folder), "quality": quality}, temporary, ensure_ascii=False)
            temporary.close()
            path.replace(SETTINGS)
        finally:
            temporary.close()
            path.unlink(missing_ok=True)


def download_command(url, folder, quality, ffmpeg, runtime):
    command = [
        str(DOWNLOADER), "-m", "yt_dlp",
        "--ignore-config", "--no-playlist", "--no-overwrites", "--windows-filenames",
        "--no-plugin-dirs", "--no-colors", "--encoding", "utf-8",
        "--ffmpeg-location", ffmpeg, "--no-js-runtimes", "--js-runtimes", runtime,
        "--socket-timeout", "20", "--retries", "3", "--fragment-retries", "3",
        "--paths", str(folder),
        "--output", "%(title).120s [%(id)s] [%(resolution)s].%(ext)s",
        "--merge-output-format", "mp4", "--remux-video", "mp4",
        "--no-simulate", "--progress", "--newline", "--progress-delta", "0.3",
        "--progress-template", "download:YTDL_PROGRESS %(progress._percent_str)s",
        "--progress-template", "postprocess:YTDL_PROCESS %(progress.status)s",
        "--print", "after_move:YTDL_FILE %(filepath)j",
    ]
    # yt-dlp sorts by the shorter dimension so the same label works for Shorts.
    resolution = "res" if quality == QUALITIES[0] else f"res:{quality.split('p', 1)[0]}"
    command += ["--format-sort", f"{resolution},vcodec:h264,acodec:aac", "--", url]
    return command


class DownloadWindow:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.process = None
        self.cancelled = False
        self.progress_running = False
        settings = load_settings()
        folder = settings.get("folder")
        quality = settings.get("quality")
        self.folder = tk.StringVar(value=folder if isinstance(folder, str) and folder else str(DEFAULT_FOLDER))
        self.quality = tk.StringVar(value=quality if quality in QUALITIES else DEFAULT_QUALITY)
        self.url = tk.StringVar()
        self.status = tk.StringVar(value="다운로드할 영상 주소를 입력하세요.")

        root.title(APP_NAME)
        root.configure(background="#f1f5f9")
        root.minsize(610, 440)
        root.option_add("*Font", "{Malgun Gothic} 10")
        root.protocol("WM_DELETE_WINDOW", self.close)

        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Card.TLabel", background="#ffffff", foreground="#17263a")
        style.configure("Muted.TLabel", background="#ffffff", foreground="#4f6177")
        style.configure("Heading.TLabel", background="#f1f5f9", foreground="#17263a", font=("Malgun Gothic", 19, "bold"))
        style.configure("TButton", padding=(12, 7))
        style.configure("Primary.TButton", background="#175cd3", foreground="white", bordercolor="#175cd3", lightcolor="#175cd3", darkcolor="#175cd3", font=("Malgun Gothic", 10, "bold"))
        style.map("Primary.TButton", background=[("disabled", "#dce5f0"), ("active", "#124ba9")], foreground=[("disabled", "#526174")])
        style.configure("Horizontal.TProgressbar", background="#175cd3", troughcolor="#e5edf6", bordercolor="#e5edf6", lightcolor="#175cd3", darkcolor="#175cd3")

        layout = ttk.Frame(root, padding=(24, 18), style="Card.TFrame")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        ttk.Label(root, text=APP_NAME, style="Heading.TLabel").grid(row=0, column=0, sticky="w", padx=24, pady=(15, 9))
        layout.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 20))
        layout.columnconfigure(0, weight=1)

        ttk.Label(layout, text="영상 주소", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        self.url_entry = ttk.Entry(layout, textvariable=self.url, width=44)
        self.url_entry.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(4, 12))
        self.paste_button = ttk.Button(layout, text="붙여넣기", command=self.paste)
        self.paste_button.grid(row=1, column=1, sticky="ew", pady=(4, 12))

        ttk.Label(layout, text="저장 폴더", style="Card.TLabel").grid(row=2, column=0, sticky="w")
        self.folder_entry = ttk.Entry(layout, textvariable=self.folder)
        self.folder_entry.grid(row=3, column=0, sticky="nsew", padx=(0, 8), pady=(4, 12))
        self.browse_button = ttk.Button(layout, text="찾아보기", command=self.browse)
        self.browse_button.grid(row=3, column=1, sticky="ew", pady=(4, 12))

        ttk.Label(layout, text="우선 화질", style="Card.TLabel").grid(row=4, column=0, sticky="w")
        quality_row = ttk.Frame(layout, style="Card.TFrame")
        quality_row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 14))
        self.quality_menu = ttk.Combobox(quality_row, textvariable=self.quality, values=QUALITIES, state="readonly", width=14)
        self.quality_menu.pack(side="left", ipady=4)
        ttk.Label(quality_row, text="원본에서 제공하는 화질에 따라 달라질 수 있습니다.", style="Muted.TLabel").pack(side="left", padx=12)

        actions = ttk.Frame(layout, style="Card.TFrame")
        actions.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.download_button = ttk.Button(actions, text="다운로드", style="Primary.TButton", command=self.start)
        self.download_button.pack(side="left")
        self.cancel_button = ttk.Button(actions, text="취소", command=self.cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=8)
        ttk.Button(actions, text="저장 폴더 열기", command=self.open_folder).pack(side="right")

        self.progress = ttk.Progressbar(layout, maximum=100)
        self.progress.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(14, 6))
        status = ttk.Label(layout, textvariable=self.status, style="Muted.TLabel", wraplength=490)
        status.grid(row=8, column=0, columnspan=2, sticky="ew")
        status.bind("<Configure>", lambda event: status.configure(wraplength=event.width))
        self.url_entry.bind("<Return>", lambda _event: self.start())
        self.url_entry.focus_set()
        root.after(100, self.read_events)

    def paste(self):
        try:
            self.url.set(self.root.clipboard_get().strip())
        except tk.TclError:
            self.status.set("클립보드에 복사한 주소가 없습니다.")
        self.url_entry.focus_set()

    def browse(self):
        initial = self.folder.get().strip()
        folder = filedialog.askdirectory(parent=self.root, initialdir=initial or str(Path.home()))
        if folder:
            self.folder.set(os.path.normpath(folder))

    def selected_folder(self):
        folder = Path(os.path.expandvars(self.folder.get().strip().strip('"'))).expanduser()
        if not folder.is_absolute():
            raise ValueError("드라이브를 포함한 전체 저장 폴더 경로를 입력하세요.")
        return folder

    def open_folder(self):
        try:
            os.startfile(self.selected_folder())
        except (OSError, ValueError) as error:
            messagebox.showerror(APP_NAME, f"저장 폴더를 열 수 없습니다.\n\n{error}", parent=self.root)

    def store_settings(self):
        try:
            save_settings(self.selected_folder(), self.quality.get())
        except (OSError, ValueError):
            return False
        return True

    def start(self):
        if self.process is not None:
            return
        try:
            url = video_url(self.url.get())
            folder = self.selected_folder()
            ffmpeg = shutil.which("ffmpeg")
            if not DOWNLOADER.is_file():
                raise OSError("다운로드 구성 요소(yt-dlp)를 찾을 수 없습니다.")
            if not ffmpeg:
                raise OSError("영상과 음성을 합치려면 FFmpeg를 설치해야 합니다.")
            runtime = "deno" if shutil.which("deno") else "node"
            if not shutil.which(runtime):
                raise OSError("유튜브 주소를 처리하려면 Node.js 또는 Deno를 설치해야 합니다.")
            folder.mkdir(parents=True, exist_ok=True)
            command = download_command(url, folder, self.quality.get(), ffmpeg, runtime)
            process = subprocess.Popen(
                command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                text=True, encoding="utf-8", errors="replace", creationflags=subprocess.CREATE_NO_WINDOW,
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
            )
        except (OSError, ValueError) as error:
            messagebox.showerror(APP_NAME, str(error), parent=self.root)
            return

        self.process = process
        self.cancelled = False
        saved = self.store_settings()
        self.set_busy(True)
        self.show_waiting("영상 정보를 확인하는 중…" if saved else "영상 정보를 확인하는 중… 설정을 저장하지 못했습니다.")
        threading.Thread(target=self.read_process, args=(process,), daemon=True).start()

    def read_process(self, process):
        filepath = None
        error = ""
        with process.stdout as output:
            for line in output:
                tag, _, value = line.strip().partition(" ")
                if tag == "YTDL_PROGRESS":
                    self.events.put(("progress", value.strip()))
                elif tag == "YTDL_PROCESS" and value == "started":
                    self.events.put(("process", None))
                elif tag == "YTDL_FILE":
                    filepath = json.loads(value)
                elif tag == "ERROR:":
                    error = value
        self.events.put(("done", (process.wait(), filepath, error)))

    def set_busy(self, busy):
        for widget in (self.url_entry, self.folder_entry, self.paste_button, self.browse_button, self.download_button):
            widget.configure(state="disabled" if busy else "normal")
        self.quality_menu.configure(state="disabled" if busy else "readonly")
        self.cancel_button.configure(state="normal" if busy else "disabled")

    def show_waiting(self, message):
        self.status.set(message)
        if not self.progress_running:
            self.progress.configure(mode="indeterminate", value=0)
            self.progress.start(13)
            self.progress_running = True

    def show_progress(self, percent):
        if self.progress_running:
            self.progress.stop()
            self.progress_running = False
        self.progress.configure(mode="determinate", value=percent)

    def read_events(self):
        while True:
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == "done":
                self.finish(*value)
            elif not self.cancelled:
                if kind == "process":
                    self.show_waiting("영상과 음성을 정리하는 중…")
                elif re.fullmatch(r"\d+(?:\.\d+)?%", value):
                    self.show_progress(float(value.rstrip("%")))
                    self.status.set(f"다운로드 중 · 현재 파일 {value}")
                else:
                    self.show_waiting("다운로드 중…")
        self.root.after(100, self.read_events)

    def finish(self, code, filepath, error):
        self.process = None
        self.set_busy(False)
        self.show_progress(0)
        if code == 0 and filepath and Path(filepath).is_file():
            self.show_progress(100)
            self.status.set(f"저장 완료 · {Path(filepath).name}")
        elif self.cancelled:
            self.status.set("다운로드를 취소했습니다.")
        else:
            self.status.set("다운로드에 실패했습니다.")
            detail = error or "저장된 영상 파일을 확인할 수 없습니다."
            messagebox.showerror(APP_NAME, f"다운로드에 실패했습니다.\n\n{detail}", parent=self.root)

    def cancel(self):
        if self.process is None or self.process.poll() is not None:
            return
        result = subprocess.run(
            ["taskkill.exe", "/PID", str(self.process.pid), "/T", "/F"],
            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            self.cancelled = True
            self.cancel_button.configure(state="disabled")
            self.status.set("다운로드를 취소하는 중…")
        elif self.process.poll() is None:
            messagebox.showerror(APP_NAME, "다운로드를 중지할 수 없습니다.", parent=self.root)

    def close(self):
        if self.process is not None:
            self.status.set("다운로드가 끝나거나 취소된 뒤 창을 닫을 수 있습니다.")
            return
        self.store_settings()
        self.root.destroy()


if __name__ == "__main__":
    window = tk.Tk()
    DownloadWindow(window)
    window.mainloop()
