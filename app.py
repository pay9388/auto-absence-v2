"""결석신고서 자동 생성기 v2 - CustomTkinter UI"""
import os
import threading
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from datetime import date as _date
from src.parser import parse_neis
from src.builder import build_hwpx

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# 다크 색상 팔레트
_BG       = '#1e1e2e'   # 창 배경
_FRAME    = '#2a2a3d'   # 카드/프레임 배경
_TREE_BG  = '#1e1e2e'   # Treeview 배경
_HEAD_BG  = '#16213e'   # Treeview 헤더
_FG       = '#cdd6f4'   # 기본 텍스트
_ACCENT   = '#89b4fa'   # 강조색 (파랑)
_SEL      = '#313244'   # Treeview 선택 배경

# 미리보기 컬럼 → student dict 매핑
PREVIEW_COLS = ['번호', '이름', '유형', '일자', '결시교시', '사유', '증빙서류']
COL_WIDTHS   = [55, 90, 110, 100, 100, 140, 110]

_FULL_TYPE = {
    ('결석', '질병'): '질병결석',
    ('결석', '인정'): '출석인정결석',
    ('지각', '인정지각'): '출석인정지각',
    ('조퇴', '인정조퇴'): '출석인정조퇴',
    ('결과', '인정결과'): '출석인정결과',
}


def _fmt_preview(student, col):
    if col == '유형':
        유형 = student.get('유형', '')
        sub  = student.get('결석종류') or student.get('조퇴종류') or ''
        return _FULL_TYPE.get((유형, sub), 유형)
    if col == '일자':
        s = student.get('시작일')
        e = student.get('종료일')
        if s and e and s != e:
            return f'{s} ~ {e}'
        return str(s) if s else ''
    if col == '결시교시':
        s, e = student.get('시작교시', ''), student.get('종료교시', '')
        return f'{s}~{e}' if s else ''
    return str(student.get(col, '') or '')


def _apply_treeview_style():
    style = ttk.Style()
    style.theme_use('default')
    style.configure('Dark.Treeview',
        background=_TREE_BG, foreground=_FG,
        fieldbackground=_TREE_BG, rowheight=26,
        borderwidth=0, font=('', 9))
    style.configure('Dark.Treeview.Heading',
        background=_HEAD_BG, foreground=_ACCENT,
        relief='flat', font=('', 9, 'bold'))
    style.map('Dark.Treeview',
        background=[('selected', _SEL)],
        foreground=[('selected', _FG)])
    style.map('Dark.Treeview.Heading',
        background=[('active', _FRAME)])
    style.configure('Dark.Scrollbar',
        background=_FRAME, troughcolor=_TREE_BG,
        arrowcolor=_FG, borderwidth=0)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('결석신고서 자동 생성기 v2')
        self.configure(fg_color=_BG)
        self.resizable(True, True)
        self.minsize(760, 560)
        self.geometry('880x640')

        self._students   = []
        self._edit_entry = None
        self._sort_col   = None
        self._sort_asc   = True
        self._filepath    = tk.StringVar()
        self._grade       = tk.StringVar(value='2')
        self._class_num   = tk.StringVar(value='2')
        self._teacher     = tk.StringVar()
        self._outpath     = tk.StringVar(
            value=os.path.join(os.path.expanduser('~'), 'Desktop', '결석신고서_출력.hwpx'))
        self._status    = tk.StringVar(value='나이스 출결 현황 파일을 선택해주세요.')
        self._count_var = tk.StringVar(value='0건')

        _apply_treeview_style()
        self._build_ui()

    def _build_ui(self):
        main = ctk.CTkFrame(self, fg_color='transparent')
        main.pack(fill='both', expand=True, padx=16, pady=12)

        # ── [1] 학교 설정 ─────────────────────────────────────────────────────
        row0 = ctk.CTkFrame(main, fg_color=_FRAME, corner_radius=10)
        row0.pack(fill='x', pady=(0, 8))

        inner0 = ctk.CTkFrame(row0, fg_color='transparent')
        inner0.pack(padx=12, pady=8)

        ctk.CTkLabel(inner0, text='학년', text_color=_FG).pack(side='left')
        ctk.CTkComboBox(inner0, values=['1', '2', '3'], variable=self._grade,
                        width=64, state='readonly',
                        fg_color=_BG, border_color=_ACCENT,
                        button_color=_ACCENT, dropdown_fg_color=_FRAME,
                        text_color=_FG).pack(side='left', padx=(4, 16))

        ctk.CTkLabel(inner0, text='반', text_color=_FG).pack(side='left')
        ctk.CTkEntry(inner0, textvariable=self._class_num, width=52,
                     fg_color=_BG, border_color=_ACCENT,
                     text_color=_FG).pack(side='left', padx=(4, 16))

        ctk.CTkLabel(inner0, text='담임교사', text_color=_FG).pack(side='left')
        ctk.CTkEntry(inner0, textvariable=self._teacher, width=110,
                     fg_color=_BG, border_color=_ACCENT,
                     text_color=_FG).pack(side='left', padx=(4, 0))

        # ── [2] 파일 선택 ─────────────────────────────────────────────────────
        row1 = ctk.CTkFrame(main, fg_color=_FRAME, corner_radius=10)
        row1.pack(fill='x', pady=(0, 8))

        inner1 = ctk.CTkFrame(row1, fg_color='transparent')
        inner1.pack(fill='x', padx=12, pady=8)

        ctk.CTkLabel(inner1, text='출결 파일', text_color=_FG).pack(side='left')
        ctk.CTkEntry(inner1, textvariable=self._filepath,
                     fg_color=_BG, border_color=_ACCENT, text_color=_FG,
                     state='disabled').pack(
            side='left', padx=(8, 8), expand=True, fill='x')
        ctk.CTkButton(inner1, text='파일 선택…', command=self._select_file,
                      width=90, corner_radius=8,
                      fg_color=_ACCENT, text_color=_BG,
                      hover_color='#74c7ec').pack(side='left')

        # ── [3] 미리보기 ─────────────────────────────────────────────────────
        preview_card = ctk.CTkFrame(main, fg_color=_FRAME, corner_radius=10)
        preview_card.pack(fill='both', expand=True, pady=(0, 8))

        # 헤더 행
        hdr = ctk.CTkFrame(preview_card, fg_color='transparent')
        hdr.pack(fill='x', padx=12, pady=(8, 4))

        ctk.CTkLabel(hdr, text='데이터 미리보기',
                     text_color=_ACCENT, font=ctk.CTkFont(size=12, weight='bold')).pack(side='left')
        ctk.CTkLabel(hdr, textvariable=self._count_var,
                     text_color=_ACCENT).pack(side='right')

        btn_row = ctk.CTkFrame(preview_card, fg_color='transparent')
        btn_row.pack(fill='x', padx=12, pady=(0, 6))
        ctk.CTkButton(btn_row, text='선택 병합', command=self._merge_selected,
                      width=84, height=28, corner_radius=6,
                      fg_color='#313244', text_color=_FG,
                      hover_color=_ACCENT).pack(side='left', padx=(0, 6))
        ctk.CTkButton(btn_row, text='병합 해제', command=self._split_selected,
                      width=84, height=28, corner_radius=6,
                      fg_color='#313244', text_color=_FG,
                      hover_color='#f38ba8').pack(side='left')

        # Treeview (tk.Frame 으로 감싸서 배경 통일)
        tree_wrap = tk.Frame(preview_card, bg=_TREE_BG)
        tree_wrap.pack(fill='both', expand=True, padx=12, pady=(0, 10))

        self._tree = ttk.Treeview(tree_wrap, columns=PREVIEW_COLS,
                                  show='headings', height=9,
                                  selectmode='extended', style='Dark.Treeview')
        for col, w in zip(PREVIEW_COLS, COL_WIDTHS):
            self._tree.heading(col, text=col,
                               command=lambda c=col: self._sort_by(c))
            self._tree.column(col, width=w, anchor='center', minwidth=40)

        vsb = ttk.Scrollbar(tree_wrap, orient='vertical',   command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_wrap, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.bind('<Double-1>', self._start_edit)

        self._tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        # ── [4] 저장 위치 ─────────────────────────────────────────────────────
        row2 = ctk.CTkFrame(main, fg_color=_FRAME, corner_radius=10)
        row2.pack(fill='x', pady=(0, 8))

        inner2 = ctk.CTkFrame(row2, fg_color='transparent')
        inner2.pack(fill='x', padx=12, pady=8)

        ctk.CTkLabel(inner2, text='저장 위치', text_color=_FG).pack(side='left')
        ctk.CTkEntry(inner2, textvariable=self._outpath,
                     fg_color=_BG, border_color=_ACCENT,
                     text_color=_FG).pack(
            side='left', padx=(8, 8), expand=True, fill='x')
        ctk.CTkButton(inner2, text='변경…', command=self._select_outpath,
                      width=70, corner_radius=8,
                      fg_color='#313244', text_color=_FG,
                      hover_color=_ACCENT).pack(side='left')

        # ── [5] 생성 버튼 + 진행 상태 ────────────────────────────────────────
        bottom_card = ctk.CTkFrame(main, fg_color=_FRAME, corner_radius=10)
        bottom_card.pack(fill='x', pady=(0, 6))

        bottom = ctk.CTkFrame(bottom_card, fg_color='transparent')
        bottom.pack(fill='x', padx=12, pady=10)

        self._gen_btn = ctk.CTkButton(
            bottom, text='  ▶  생성하기', command=self._generate,
            width=130, height=36, corner_radius=10,
            fg_color=_ACCENT, text_color=_BG,
            hover_color='#74c7ec', font=ctk.CTkFont(size=13, weight='bold'))
        self._gen_btn.pack(side='left', padx=(0, 14))

        progress_frame = ctk.CTkFrame(bottom, fg_color='transparent')
        progress_frame.pack(side='left', fill='x', expand=True)

        self._progress = ctk.CTkProgressBar(progress_frame,
                                            fg_color='#313244',
                                            progress_color=_ACCENT)
        self._progress.set(0)
        self._progress.pack(fill='x', pady=(2, 4))

        ctk.CTkLabel(progress_frame, textvariable=self._status,
                     text_color='#6c7086', anchor='w').pack(anchor='w')

        # ── [6] 하단 푸터 ────────────────────────────────────────────────────
        ctk.CTkFrame(main, height=1, fg_color='#313244',
                     corner_radius=0).pack(fill='x', pady=(4, 4))
        footer = ctk.CTkFrame(main, fg_color='transparent')
        footer.pack(fill='x')

        ctk.CTkLabel(footer, text='© 2026 TeacherCHO84',
                     text_color='#a6adc8', font=ctk.CTkFont(size=14)).pack(side='left')
        email_lbl = ctk.CTkLabel(footer, text='teachercho84@gmail.com',
                                 text_color=_ACCENT, font=ctk.CTkFont(size=14),
                                 cursor='hand2')
        email_lbl.pack(side='right')
        email_lbl.bind('<Button-1>',
                       lambda e: webbrowser.open('mailto:teachercho84@gmail.com'))

    # ── 병합 / 해제 ───────────────────────────────────────────────────────────

    def _merge_selected(self):
        selected = self._tree.selection()
        if len(selected) < 2:
            messagebox.showwarning('경고', '2개 이상 선택해주세요.')
            return
        indices = sorted(self._tree.index(iid) for iid in selected)
        students = [self._students[i] for i in indices]

        if len(set(s['이름'] for s in students)) > 1:
            messagebox.showwarning('경고', '같은 학생의 항목만 병합할 수 있습니다.')
            return
        if len(set((s['유형'], s['결석종류']) for s in students)) > 1:
            messagebox.showwarning('경고', '같은 유형의 항목만 병합할 수 있습니다.')
            return

        sorted_s = sorted(students, key=lambda s: s['시작일'])
        merged = dict(sorted_s[0])
        merged['종료일'] = sorted_s[-1]['시작일']
        merged['_원본'] = [dict(s) for s in sorted_s]

        for i in reversed(indices):
            self._students.pop(i)
        self._students.insert(indices[0], merged)
        self._refresh_tree()
        self._count_var.set(f'{len(self._students)}건 로드됨')

    def _split_selected(self):
        selected = self._tree.selection()
        if len(selected) != 1:
            messagebox.showwarning('경고', '해제할 항목을 1개 선택해주세요.')
            return
        iid = selected[0]
        idx = self._tree.index(iid)
        s = self._students[idx]
        originals = s.get('_원본')
        if not originals:
            messagebox.showinfo('안내', '병합된 항목이 아닙니다.')
            return

        self._students.pop(idx)
        for i, orig in enumerate(originals):
            self._students.insert(idx + i, orig)
        self._refresh_tree()
        self._count_var.set(f'{len(self._students)}건 로드됨')

    # ── 정렬 ──────────────────────────────────────────────────────────────────

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True

        def _key(s):
            val = _fmt_preview(s, col)
            if col == '번호':
                try:
                    return int(val)
                except (ValueError, TypeError):
                    return 0
            return val

        self._students.sort(key=_key, reverse=not self._sort_asc)
        self._refresh_tree()

        arrow = ' ▲' if self._sort_asc else ' ▼'
        for c in PREVIEW_COLS:
            self._tree.heading(c, text=c + (arrow if c == col else ''))

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for s in self._students:
            self._tree.insert('', 'end', values=[_fmt_preview(s, c) for c in PREVIEW_COLS])

    # ── 인라인 셀 편집 ────────────────────────────────────────────────────────

    def _start_edit(self, event):
        if self._tree.identify_region(event.x, event.y) != 'cell':
            return
        iid = self._tree.identify_row(event.y)
        col = self._tree.identify_column(event.x)
        if not iid or not col:
            return

        col_idx  = int(col[1:]) - 1
        col_name = PREVIEW_COLS[col_idx]
        bbox     = self._tree.bbox(iid, col)
        if not bbox:
            return
        x, y, width, height = bbox
        current = self._tree.item(iid, 'values')[col_idx]

        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None

        var = tk.StringVar(value=current)

        if col_name == '증빙서류':
            _JEUNGBI = ['', '진단서', '소견서', '진료 확인서', '처방전',
                        '공문', '담임교사 확인서', '기타']
            entry = ttk.Combobox(self._tree, textvariable=var,
                                 values=_JEUNGBI, state='readonly',
                                 font=('', 9))
            entry.place(x=x, y=y, width=width, height=height)
            entry.focus_set()
        else:
            entry = tk.Entry(self._tree, textvariable=var,
                             bg='#313244', fg=_FG,
                             insertbackground=_FG,
                             relief='flat', font=('', 9))
            entry.place(x=x, y=y, width=width, height=height)
            entry.focus_set()
            entry.select_range(0, 'end')
        self._edit_entry = entry

        def _save(event=None):
            if not self._edit_entry:
                return
            new_val = var.get().strip()
            vals = list(self._tree.item(iid, 'values'))
            vals[col_idx] = new_val
            self._tree.item(iid, values=vals)

            row_idx = self._tree.index(iid)
            if row_idx < len(self._students):
                s = self._students[row_idx]
                if col_name == '일자':
                    try:
                        parts = new_val.split('-')
                        d = _date(int(parts[0]), int(parts[1]), int(parts[2]))
                        s['시작일'] = d
                        s['종료일'] = d
                    except Exception:
                        pass
                elif col_name == '결시교시':
                    if '~' in new_val:
                        parts = new_val.split('~', 1)
                        s['시작교시'] = parts[0].strip()
                        s['종료교시'] = parts[1].strip()
                    elif new_val:
                        s['시작교시'] = new_val
                        s['종료교시'] = new_val
                else:
                    s[col_name] = new_val

            entry.destroy()
            self._edit_entry = None

        def _cancel(event=None):
            if self._edit_entry:
                entry.destroy()
                self._edit_entry = None

        entry.bind('<Return>',   _save)
        entry.bind('<Tab>',      _save)
        entry.bind('<Escape>',   _cancel)
        entry.bind('<FocusOut>', _save)

    # ── 이벤트 핸들러 ─────────────────────────────────────────────────────────

    def _select_file(self):
        path = filedialog.askopenfilename(
            title='나이스 출결 현황 파일 선택',
            filetypes=[('Excel 파일', '*.xlsx *.xls *.xlsm'), ('All files', '*.*')]
        )
        if not path:
            return
        self._filepath.set(path)
        self._load_preview(path)

    def _load_preview(self, path):
        try:
            students = parse_neis(path)
            self._students = students
            self._refresh_tree()
            self._count_var.set(f'{len(students)}명 로드됨')
            self._status.set(f'{len(students)}명 데이터 로드 완료.')
        except Exception as e:
            messagebox.showerror('파일 오류', str(e))
            self._status.set('파일 로드 실패.')

    def _select_outpath(self):
        path = filedialog.asksaveasfilename(
            title='저장 위치 선택',
            defaultextension='.hwpx',
            filetypes=[('HWPX 파일', '*.hwpx'), ('All files', '*.*')]
        )
        if path:
            self._outpath.set(path)

    def _generate(self):
        if not self._students:
            messagebox.showwarning('경고', '먼저 출결 파일을 선택해주세요.')
            return
        outpath = self._outpath.get().strip()
        if not outpath:
            messagebox.showwarning('경고', '저장 위치를 지정해주세요.')
            return
        grade     = self._grade.get().strip()
        class_num = self._class_num.get().strip()
        if not class_num.isdigit():
            messagebox.showwarning('경고', '반 번호를 올바르게 입력해주세요.')
            return
        teacher_name = self._teacher.get().strip()

        self._gen_btn.configure(state='disabled')
        self._progress.set(0)
        self._status.set('생성 중…')

        def _run():
            try:
                def on_progress(i, total):
                    pct = i / total if total > 0 else 0
                    self.after(0, lambda p=pct, ci=i, t=total:
                               (self._progress.set(p),
                                self._status.set(f'처리 중… {ci}/{t}명')))

                build_hwpx(self._students, grade, class_num, teacher_name,
                           outpath, on_progress)
                self.after(0, self._on_done, outpath)
            except Exception as e:
                self.after(0, lambda err=str(e):
                           (messagebox.showerror('생성 실패', err),
                            self._status.set('오류 발생.'),
                            self._gen_btn.configure(state='normal')))

        threading.Thread(target=_run, daemon=True).start()

    def _on_done(self, outpath):
        self._progress.set(1)
        self._status.set(f'완료: {os.path.basename(outpath)}')
        self._gen_btn.configure(state='normal')
        if messagebox.askyesno(
                '생성 완료',
                f'결석신고서 {len(self._students)}명분 생성 완료!\n\n'
                f'{outpath}\n\n파일을 지금 열어볼까요?'):
            os.startfile(outpath)


# ── 진입점 ──────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = App()
    app.mainloop()
