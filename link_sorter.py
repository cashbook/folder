"""
링크 파일 자동 분류 프로그램
- 링크 파일(.url, .lnk 등)을 드래그 앤 드롭하면
- Gemini AI가 파일 이름을 분석하여
- D:\06. 기타 하위 폴더 중 가장 적합한 곳으로 자동 이동
"""

import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from google import genai

# ── 설정 ──
BASE_DIR = r"D:\06. 기타"
DIRECT_FOLDER = "볼 것"
GEMINI_API_KEY = "AIzaSyCr6x7piW3xInaM8nI7H0OYydl9_Oip2C8"
STARTUP_DIR = os.path.join(
    os.environ["APPDATA"],
    r"Microsoft\Windows\Start Menu\Programs\Startup",
)
STARTUP_LNK = os.path.join(STARTUP_DIR, "LinkSorter.lnk")
SCRIPT_PATH = os.path.abspath(__file__)

client = genai.Client(api_key=GEMINI_API_KEY)


def get_subfolders(base_dir):
    """BASE_DIR의 1단계 하위 폴더 목록을 가져온다."""
    folders = []
    try:
        for entry in os.scandir(base_dir):
            if entry.is_dir():
                folders.append(entry.name)
    except FileNotFoundError:
        messagebox.showerror("오류", f"폴더를 찾을 수 없습니다: {base_dir}")
    return sorted(folders)


def ask_gemini(file_name, folder_list):
    """Gemini API에 파일 이름과 폴더 목록을 보내 적합한 폴더를 판단한다."""
    prompt = f"""당신은 파일 분류 전문가입니다.
아래 파일 이름을 보고, 주어진 폴더 목록 중 가장 적합한 폴더 이름을 **정확히 하나만** 골라 답하세요.
답변은 폴더 이름만 출력하세요. 부연 설명 없이 폴더 이름만 적으세요.

반드시 지켜야 할 규칙:
- 파일 이름에 인간을 제외한 동물(개, 고양이, 강아지, 새, 물고기, 곤충, 파충류, 포유류, 야생동물 등 모든 동물)이 언급되어 있으면 반드시 "동물" 폴더로 분류하세요. 이 규칙은 다른 판단보다 우선합니다.

파일 이름: {file_name}

폴더 목록:
{chr(10).join(folder_list)}
"""
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=prompt,
    )
    chosen = response.text.strip()
    # 혹시 줄바꿈이나 공백이 포함되었으면 첫 줄만
    chosen = chosen.splitlines()[0].strip()
    return chosen


def parse_dropped_paths(data):
    """tkinterdnd2가 넘겨주는 문자열을 파일 경로 리스트로 변환한다."""
    paths = []
    current = ""
    in_brace = False
    for ch in data:
        if ch == '{':
            in_brace = True
            continue
        elif ch == '}':
            in_brace = False
            if current:
                paths.append(current)
                current = ""
            continue
        elif ch == ' ' and not in_brace:
            if current:
                paths.append(current)
                current = ""
            continue
        current += ch
    if current:
        paths.append(current)
    return paths


def move_file(src, dest_dir):
    """파일을 dest_dir로 이동한다. 중복 시 번호를 붙인다. (dest 경로 반환)"""
    file_name = os.path.basename(src)
    dest_path = os.path.join(dest_dir, file_name)
    if os.path.exists(dest_path):
        name, ext = os.path.splitext(file_name)
        counter = 1
        while os.path.exists(dest_path):
            dest_path = os.path.join(dest_dir, f"{name} ({counter}){ext}")
            counter += 1
    shutil.move(src, dest_path)
    return dest_path


def is_autostart_enabled():
    """시작 프로그램 바로가기가 존재하는지 확인한다."""
    return os.path.exists(STARTUP_LNK)


def set_autostart(enable):
    """시작 프로그램 등록/해제."""
    if enable:
        pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if not os.path.exists(pythonw):
            pythonw = sys.executable
        ps_script = (
            f'$ws = New-Object -ComObject WScript.Shell; '
            f'$sc = $ws.CreateShortcut("{STARTUP_LNK}"); '
            f'$sc.TargetPath = "{pythonw}"; '
            f'$sc.Arguments = \'"{SCRIPT_PATH}"\'; '
            f'$sc.WorkingDirectory = "{os.path.dirname(SCRIPT_PATH)}"; '
            f'$sc.Save()'
        )
        subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    else:
        if os.path.exists(STARTUP_LNK):
            os.remove(STARTUP_LNK)


class LinkSorterApp:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("링크 파일 자동 분류")
        self.root.geometry("700x580")
        self.root.configure(bg="#1e1e2e")
        self.root.resizable(False, False)

        self.folders = get_subfolders(BASE_DIR)

        self._build_ui()

    def _build_ui(self):
        # ── AI 분류 드롭 영역 (왼쪽) ──
        self.drop_ai = tk.Frame(
            self.root, bg="#313244", highlightbackground="#89b4fa",
            highlightthickness=2, cursor="hand2"
        )
        self.drop_ai.place(x=20, y=20, width=400, height=160)

        self.ai_label = tk.Label(
            self.drop_ai,
            text="AI 자동 분류",
            font=("맑은 고딕", 14, "bold"),
            fg="#89b4fa", bg="#313244"
        )
        self.ai_label.place(relx=0.5, rely=0.3, anchor="center")

        self.ai_sub = tk.Label(
            self.drop_ai,
            text="Gemini가 알맞은 폴더를 골라줍니다",
            font=("맑은 고딕", 9),
            fg="#6c7086", bg="#313244"
        )
        self.ai_sub.place(relx=0.5, rely=0.55, anchor="center")

        self.ai_ext = tk.Label(
            self.drop_ai,
            text=".url  .lnk  또는 기타 파일",
            font=("맑은 고딕", 9),
            fg="#45475a", bg="#313244"
        )
        self.ai_ext.place(relx=0.5, rely=0.75, anchor="center")

        self.drop_ai.drop_target_register(DND_FILES)
        self.drop_ai.dnd_bind("<<Drop>>", self._on_drop_ai)
        self.drop_ai.dnd_bind("<<DragEnter>>", lambda e: self._highlight(self.drop_ai, self.ai_label, True))
        self.drop_ai.dnd_bind("<<DragLeave>>", lambda e: self._highlight(self.drop_ai, self.ai_label, False, "AI 자동 분류", "#89b4fa"))

        # ── '볼 것' 직접 이동 드롭 영역 (오른쪽) ──
        self.drop_direct = tk.Frame(
            self.root, bg="#313244", highlightbackground="#f9e2af",
            highlightthickness=2, cursor="hand2"
        )
        self.drop_direct.place(x=440, y=20, width=240, height=160)

        self.direct_label = tk.Label(
            self.drop_direct,
            text="볼 것",
            font=("맑은 고딕", 16, "bold"),
            fg="#f9e2af", bg="#313244"
        )
        self.direct_label.place(relx=0.5, rely=0.35, anchor="center")

        self.direct_sub = tk.Label(
            self.drop_direct,
            text=f"→ {DIRECT_FOLDER} 폴더로 바로 이동",
            font=("맑은 고딕", 9),
            fg="#6c7086", bg="#313244"
        )
        self.direct_sub.place(relx=0.5, rely=0.65, anchor="center")

        self.drop_direct.drop_target_register(DND_FILES)
        self.drop_direct.dnd_bind("<<Drop>>", self._on_drop_direct)
        self.drop_direct.dnd_bind("<<DragEnter>>", lambda e: self._highlight(self.drop_direct, self.direct_label, True))
        self.drop_direct.dnd_bind("<<DragLeave>>", lambda e: self._highlight(self.drop_direct, self.direct_label, False, "볼 것", "#f9e2af"))

        # ── 부팅 시 자동 실행 체크박스 ──
        self.autostart_var = tk.BooleanVar(value=is_autostart_enabled())
        self.autostart_cb = tk.Checkbutton(
            self.root, text="부팅 시 자동 실행",
            variable=self.autostart_var,
            command=self._toggle_autostart,
            font=("맑은 고딕", 10),
            fg="#a6adc8", bg="#1e1e2e",
            selectcolor="#313244", activebackground="#1e1e2e",
            activeforeground="#cdd6f4",
        )
        self.autostart_cb.place(x=20, y=192)

        # ── 로그 영역 ──
        log_label = tk.Label(
            self.root, text="처리 로그",
            font=("맑은 고딕", 11, "bold"),
            fg="#a6adc8", bg="#1e1e2e"
        )
        log_label.place(x=20, y=225)

        self.log_text = tk.Text(
            self.root, font=("Consolas", 10),
            fg="#cdd6f4", bg="#181825",
            insertbackground="#cdd6f4",
            selectbackground="#45475a",
            relief="flat", wrap="word"
        )
        self.log_text.place(x=20, y=255, width=660, height=300)
        self.log_text.config(state="disabled")

        # 폴더 수 표시
        info = tk.Label(
            self.root, text=f"대상 폴더: {BASE_DIR}  |  하위 폴더 {len(self.folders)}개 감지됨",
            font=("맑은 고딕", 9),
            fg="#585b70", bg="#1e1e2e"
        )
        info.place(x=20, y=558)

    # ── 드래그 시각 피드백 ──
    def _highlight(self, frame, label, enter, text=None, color=None):
        if enter:
            frame.config(highlightbackground="#a6e3a1", highlightthickness=3)
            label.config(text="여기에 놓으세요!", fg="#a6e3a1")
        else:
            frame.config(highlightbackground=color or "#89b4fa", highlightthickness=2)
            label.config(text=text or "", fg=color or "#cdd6f4")

    # ── AI 분류 드롭 ──
    def _on_drop_ai(self, event):
        self._highlight(self.drop_ai, self.ai_label, False, "AI 자동 분류", "#89b4fa")
        paths = parse_dropped_paths(event.data)
        if paths:
            threading.Thread(target=self._process_ai, args=(paths,), daemon=True).start()

    def _process_ai(self, paths):
        for path in paths:
            if not os.path.isfile(path):
                self._log(f"[건너뜀] 파일이 아님: {path}")
                continue

            file_name = os.path.basename(path)
            self._log(f"[분석 중] {file_name}")

            try:
                chosen = ask_gemini(file_name, self.folders)
            except Exception as e:
                self._log(f"  [오류] Gemini API 호출 실패: {e}")
                continue

            dest_dir = os.path.join(BASE_DIR, chosen)
            if not os.path.isdir(dest_dir):
                self._log(f"  [오류] AI가 선택한 폴더가 존재하지 않음: {chosen}")
                match = self._find_closest_folder(chosen)
                if match:
                    self._log(f"  [보정] 유사 폴더 발견: {match}")
                    chosen = match
                    dest_dir = os.path.join(BASE_DIR, chosen)
                else:
                    self._log(f"  [실패] 적합한 폴더를 찾지 못했습니다.")
                    continue

            try:
                dest = move_file(path, dest_dir)
                self._log(f"  [완료] → {chosen}/{os.path.basename(dest)}")
            except Exception as e:
                self._log(f"  [오류] 파일 이동 실패: {e}")

    # ── '볼 것' 직접 드롭 ──
    def _on_drop_direct(self, event):
        self._highlight(self.drop_direct, self.direct_label, False, "볼 것", "#f9e2af")
        paths = parse_dropped_paths(event.data)
        if paths:
            threading.Thread(target=self._process_direct, args=(paths,), daemon=True).start()

    def _process_direct(self, paths):
        dest_dir = os.path.join(BASE_DIR, DIRECT_FOLDER)
        if not os.path.isdir(dest_dir):
            self._log(f"[오류] 폴더가 존재하지 않음: {dest_dir}")
            return

        for path in paths:
            if not os.path.isfile(path):
                self._log(f"[건너뜀] 파일이 아님: {path}")
                continue

            file_name = os.path.basename(path)
            try:
                dest = move_file(path, dest_dir)
                self._log(f"[볼 것] {file_name} → {DIRECT_FOLDER}/")
            except Exception as e:
                self._log(f"[오류] {file_name} 이동 실패: {e}")

    # ── 자동 실행 토글 ──
    def _toggle_autostart(self):
        enable = self.autostart_var.get()
        try:
            set_autostart(enable)
            state = "등록" if enable else "해제"
            self._log(f"[설정] 부팅 시 자동 실행 {state}됨")
        except Exception as e:
            self._log(f"[오류] 자동 실행 설정 실패: {e}")
            self.autostart_var.set(not enable)

    # ── 유틸 ──
    def _find_closest_folder(self, chosen):
        chosen_lower = chosen.lower()
        for f in self.folders:
            if f.lower() == chosen_lower:
                return f
        for f in self.folders:
            if chosen_lower in f.lower() or f.lower() in chosen_lower:
                return f
        return None

    def _log(self, message):
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _append)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = LinkSorterApp()
    app.run()
