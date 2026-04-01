"""결석신고서 자동 생성기 v2 - Tkinter UI"""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from datetime import date as _date
from src.parser import parse_neis
from src.builder import build_hwpx


# 미리보기 컬럼 → student dict 매핑
PREVIEW_COLS = ['번호', '이름', '유형', '일자', '결시교시', '사유', '증빙서류']
COL_WIDTHS   = [55, 90, 65, 100, 100, 140, 110]


def _fmt_preview(student, col):
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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('결석신고서 자동 생성기 v2')
        self.resizable(True, True)
        self.minsize(760, 540)
        self.geometry('860x600')

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
        self._count_var = tk.StringVar(value='0명')

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except Exception:
            pass

        main = ttk.Frame(self, padding=10)
        main.pack(fill='both', expand=True)

        # ── [1] 학교 설정 ─────────────────────────────────────────────────────
        row0 = ttk.Frame(main)
        row0.pack(fill='x', pady=(0, 4))

        ttk.Label(row0, text='학년:').pack(side='left')
        ttk.Combobox(row0, textvariable=self._grade,
                     values=['1', '2', '3'], width=3,
                     state='readonly').pack(side='left', padx=(2, 12))

        ttk.Label(row0, text='반:').pack(side='left')
        ttk.Entry(row0, textvariable=self._class_num, width=4).pack(side='left', padx=(2, 12))

        ttk.Label(row0, text='담임교사:').pack(side='left')
        ttk.Entry(row0, textvariable=self._teacher, width=10).pack(side='left', padx=(2, 0))

        # ── [2] 파일 선택 ─────────────────────────────────────────────────────
        row1 = ttk.Frame(main)
        row1.pack(fill='x', pady=(0, 6))

        ttk.Label(row1, text='출결 파일:').pack(side='left')
        ttk.Entry(row1, textvariable=self._filepath, width=46,
                  state='readonly').pack(side='left', padx=(4, 4))
        ttk.Button(row1, text='파일 선택…',
                   command=self._select_file).pack(side='left', padx=2)

        # ── [3] 미리보기 ─────────────────────────────────────────────────────
        preview_frame = ttk.LabelFrame(main, text='데이터 미리보기', padding=4)
        preview_frame.pack(fill='both', expand=True, pady=(0, 6))

        count_row = ttk.Frame(preview_frame)
        count_row.pack(fill='x')
        ttk.Button(count_row, text='선택 병합',
                   command=self._merge_selected).pack(side='left', padx=(0, 6))
        ttk.Button(count_row, text='병합 해제',
                   command=self._split_selected).pack(side='left')
        ttk.Label(count_row, textvariable=self._count_var,
                  foreground='#1a6bbf').pack(side='right')

        tree_wrap = ttk.Frame(preview_frame)
        tree_wrap.pack(fill='both', expand=True)

        self._tree = ttk.Treeview(tree_wrap, columns=PREVIEW_COLS,
                                  show='headings', height=9,
                                  selectmode='extended')
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
        row2 = ttk.Frame(main)
        row2.pack(fill='x', pady=(0, 6))

        ttk.Label(row2, text='저장 위치:').pack(side='left')
        ttk.Entry(row2, textvariable=self._outpath, width=52).pack(
            side='left', padx=(4, 4), expand=True, fill='x')
        ttk.Button(row2, text='변경…',
                   command=self._select_outpath).pack(side='left', padx=2)

        # ── [5] 생성 버튼 + 진행 상태 ────────────────────────────────────────
        bottom = ttk.Frame(main)
        bottom.pack(fill='x')

        self._gen_btn = ttk.Button(
            bottom, text='  ▶  생성하기  ', command=self._generate, width=18)
        self._gen_btn.pack(side='left', padx=(0, 10))

        progress_frame = ttk.Frame(bottom)
        progress_frame.pack(side='left', fill='x', expand=True)

        self._progress = ttk.Progressbar(progress_frame, mode='determinate', length=300)
        self._progress.pack(fill='x')
        ttk.Label(progress_frame, textvariable=self._status,
                  foreground='gray').pack(anchor='w')

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
        # 원본 날짜 목록 저장 (병합 해제용)
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

        var   = tk.StringVar(value=current)
        entry = ttk.Entry(self._tree, textvariable=var)
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
                    # "X~Y" 형식
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

        self._gen_btn.config(state='disabled')
        self._progress['value'] = 0
        self._status.set('생성 중…')

        def _run():
            try:
                def on_progress(i, total):
                    pct = int(i / total * 100) if total > 0 else 0
                    self.after(0, lambda p=pct, ci=i, t=total:
                               (self._progress.config(value=p),
                                self._status.set(f'처리 중… {ci}/{t}명')))

                build_hwpx(self._students, grade, class_num, teacher_name,
                           outpath, on_progress)
                self.after(0, self._on_done, outpath)
            except Exception as e:
                self.after(0, lambda err=str(e):
                           (messagebox.showerror('생성 실패', err),
                            self._status.set('오류 발생.'),
                            self._gen_btn.config(state='normal')))

        threading.Thread(target=_run, daemon=True).start()

    def _on_done(self, outpath):
        self._progress['value'] = 100
        self._status.set(f'완료: {os.path.basename(outpath)}')
        self._gen_btn.config(state='normal')
        if messagebox.askyesno(
                '생성 완료',
                f'결석신고서 {len(self._students)}명분 생성 완료!\n\n'
                f'{outpath}\n\n파일을 지금 열어볼까요?'):
            os.startfile(outpath)


# ── 진입점 ──────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = App()
    app.mainloop()
