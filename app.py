#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ناسخ (Nasikh) — برنامج البحث في "الجامع" ونسخ النتائج إلى Word
==============================================================

برنامج جاهز بواجهة رسومية: تبحث عن عبارة، تحدّد النتائج التي تريدها،
وتضغط زر «نسخ إلى Word» فتُنشأ وثيقة Word (.docx) فيها النصوص المحددة
وكل نصّ متبوع بتخريجه (اسم الكتاب، المؤلف، رقم الصفحة) + رابط المصدر.

نسخ النتائج إلى Word هو الغاية الكبرى للبرنامج، وهي غير متوفرة في المواقع.

التشغيل:
    1) ثبّت بايثون من python.org (مع تحديد "Add to PATH").
    2) ثبّت المكتبتين:   pip install requests python-docx
    3) شغّل:            python app.py

(على ويندوز وماك، Tkinter مرفق مع بايثون. على لينكس: sudo apt install python3-tk)
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import aljam3


class TakhrijApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ناسخ — البحث في الجامع ونسخ النتائج إلى Word")
        self.geometry("1150x760")
        self.minsize(960, 640)

        self.hits: list[aljam3.SearchHit] = []
        self.checked: dict[str, aljam3.SearchHit] = {}
        self.pagination: dict = {}
        self.categories: dict[int, str] = {0: "— كل الأقسام —"}

        self._build_search_bar()
        self._build_results()
        self._build_preview()
        self._build_statusbar()

        # حمّل الأقسام في الخلفية
        threading.Thread(target=self._load_categories, daemon=True).start()

    # ----------------------------------------------------------------- بناء الواجهة
    def _build_search_bar(self) -> None:
        frm = ttk.LabelFrame(self, text="البحث")
        frm.pack(fill="x", padx=8, pady=6)

        row1 = ttk.Frame(frm)
        row1.pack(fill="x", padx=6, pady=5)
        ttk.Label(row1, text="العبارة:").pack(side="right")
        self.q_var = tk.StringVar()
        e = ttk.Entry(row1, textvariable=self.q_var, justify="right", font=("Arial", 12))
        e.pack(side="right", fill="x", expand=True, padx=6)
        e.bind("<Return>", lambda ev: self.do_search())
        ttk.Button(row1, text="🔎 ابحث", command=self.do_search).pack(side="right")

        row2 = ttk.Frame(frm)
        row2.pack(fill="x", padx=6, pady=5)

        ttk.Label(row2, text="القسم:").pack(side="right")
        self.cat_var = tk.StringVar(value="— كل الأقسام —")
        self.cat_cb = ttk.Combobox(
            row2, textvariable=self.cat_var, state="readonly", width=26,
            justify="right", values=["— كل الأقسام —"],
        )
        self.cat_cb.pack(side="right", padx=(2, 12))

        ttk.Label(row2, text="رقم الكتاب (اختياري):").pack(side="right")
        self.book_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.book_var, width=8, justify="right").pack(
            side="right", padx=(2, 12)
        )

        ttk.Label(row2, text="عدد النتائج:").pack(side="right")
        self.limit_var = tk.IntVar(value=20)
        ttk.Spinbox(row2, from_=5, to=100, textvariable=self.limit_var, width=5).pack(
            side="right", padx=(2, 12)
        )

    def _build_results(self) -> None:
        frm = ttk.LabelFrame(
            self, text="النتائج  (نقر مزدوج أو مسطرة المسافة = تحديد/إلغاء)"
        )
        frm.pack(fill="both", expand=True, padx=8, pady=6)

        cols = ("sel", "snippet", "book", "author", "page")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", height=11)
        for c, (txt, w, anchor) in {
            "sel": ("✔", 36, "center"),
            "snippet": ("المقتطف", 470, "e"),
            "book": ("الكتاب", 230, "e"),
            "author": ("المؤلف", 180, "e"),
            "page": ("الصفحة", 60, "center"),
        }.items():
            self.tree.heading(c, text=txt)
            self.tree.column(c, width=w, anchor=anchor)

        vsb = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_row_focus)
        self.tree.bind("<space>", self._toggle)
        self.tree.bind("<Double-1>", self._toggle)

        # شريط ترقيم الصفحات + تحديد الكل
        nav = ttk.Frame(frm)
        nav.pack(side="bottom", fill="x")

    def _build_preview(self) -> None:
        frm = ttk.LabelFrame(self, text="معاينة النص المحدَّد")
        frm.pack(fill="both", expand=True, padx=8, pady=6)

        self.preview = tk.Text(frm, height=7, wrap="word", font=("Arial", 13))
        self.preview.tag_configure("rtl", justify="right")
        self.preview.pack(fill="both", expand=True, padx=6, pady=4)

        self.cit_var = tk.StringVar(value="التخريج سيظهر هنا بعد اختيار نتيجة…")
        ttk.Label(frm, textvariable=self.cit_var, foreground="#065").pack(fill="x", padx=6)

        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=6, pady=8)

        # الزر البطل
        self.with_author_var = tk.BooleanVar(value=True)
        self.with_link_var = tk.BooleanVar(value=True)

        export_btn = ttk.Button(
            bar, text="📄  نسخ المحدَّد إلى Word", command=self.export_word
        )
        export_btn.pack(side="left", padx=4, ipadx=10, ipady=4)

        ttk.Button(bar, text="تحديد الكل", command=self._select_all).pack(side="left", padx=4)
        ttk.Button(bar, text="إلغاء التحديد", command=self._clear_all).pack(side="left", padx=4)

        ttk.Checkbutton(bar, text="رابط المصدر في Word", variable=self.with_link_var).pack(
            side="right", padx=6
        )
        ttk.Checkbutton(
            bar, text="إدراج المؤلف في التخريج", variable=self.with_author_var,
            command=self._refresh_cit,
        ).pack(side="right", padx=6)

    def _build_statusbar(self) -> None:
        self.status = tk.StringVar(value="جاهز. اكتب عبارة واضغط ابحث.")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="e").pack(
            fill="x", side="bottom"
        )

    # ----------------------------------------------------------------- مساعدات
    def _set_status(self, msg: str) -> None:
        self.status.set(msg)
        self.update_idletasks()

    def _load_categories(self) -> None:
        try:
            cats = aljam3.list_categories()
        except Exception:  # noqa: BLE001
            return
        self.categories = {0: "— كل الأقسام —"}
        for c in cats:
            self.categories[c.get("id")] = aljam3.clean_text(c.get("name", ""))
        self.after(0, lambda: self.cat_cb.configure(values=list(self.categories.values())))

    # ----------------------------------------------------------------- البحث
    def do_search(self) -> None:
        q = self.q_var.get().strip()
        if not q:
            messagebox.showwarning("تنبيه", "اكتب عبارة البحث أولًا.")
            return
        cat_id = next((k for k, v in self.categories.items() if v == self.cat_var.get()), 0)
        book = self.book_var.get().strip()
        books = [int(book)] if book.isdigit() else None
        cats = [cat_id] if cat_id else None
        limit = self.limit_var.get()

        self._set_status("جارٍ البحث في الجامع…")
        threading.Thread(
            target=self._search_worker, args=(q, books, cats, limit), daemon=True
        ).start()

    def _search_worker(self, q, books, cats, limit) -> None:
        try:
            pagination, hits = aljam3.search(
                q, books=books, categories=cats, limit=limit, enrich=True
            )
        except Exception as e:  # noqa: BLE001
            self.after(0, lambda: self._search_err(e))
            return
        self.after(0, lambda: self._search_done(pagination, hits))

    def _search_err(self, e: Exception) -> None:
        self._set_status("فشل الاتصال.")
        messagebox.showerror(
            "خطأ في الاتصال",
            f"تعذّر الاتصال بالجامع (aljam3.com):\n{e}\n\n"
            "تأكد من اتصال الإنترنت.",
        )

    def _search_done(self, pagination, hits) -> None:
        self.pagination = pagination
        self.hits = hits
        self.checked.clear()
        self.tree.delete(*self.tree.get_children())
        for i, hit in enumerate(hits):
            snip = aljam3.clean_text(hit.content)
            self.tree.insert(
                "", "end", iid=str(i),
                values=(
                    "☐",
                    snip[:130] + ("…" if len(snip) > 130 else ""),
                    hit.book.title if hit.book else "—",
                    hit.book.author_name if hit.book else "—",
                    hit.page_number,
                ),
            )
        self._set_status(
            f"إجمالي النتائج: {pagination.get('count', '؟')} — المعروض: {len(hits)}"
        )

    # ----------------------------------------------------------------- التحديد
    def _toggle(self, event=None) -> str:
        item = self.tree.focus()
        if not item:
            return "break"
        if item in self.checked:
            del self.checked[item]
            self.tree.set(item, "sel", "☐")
        else:
            self.checked[item] = self.hits[int(item)]
            self.tree.set(item, "sel", "☑")
        self._set_status(f"المحدَّد للنسخ إلى Word: {len(self.checked)}")
        return "break"

    def _select_all(self) -> None:
        for i, hit in enumerate(self.hits):
            self.checked[str(i)] = hit
            self.tree.set(str(i), "sel", "☑")
        self._set_status(f"المحدَّد للنسخ إلى Word: {len(self.checked)}")

    def _clear_all(self) -> None:
        self.checked.clear()
        for i in range(len(self.hits)):
            self.tree.set(str(i), "sel", "☐")
        self._set_status("أُلغي التحديد.")

    def _on_row_focus(self, event=None) -> None:
        item = self.tree.focus()
        if not item:
            return
        hit = self.hits[int(item)]
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", aljam3.clean_text(hit.content), "rtl")
        self._current_hit = hit
        self._refresh_cit()

    def _refresh_cit(self) -> None:
        hit = getattr(self, "_current_hit", None)
        if not hit:
            return
        cit = aljam3.format_citation(hit, with_author=self.with_author_var.get())
        self.cit_var.set(f"التخريج:  ({cit})")

    # ----------------------------------------------------------------- النقل إلى Word
    def export_word(self) -> None:
        if not self.checked:
            messagebox.showinfo(
                "لا يوجد تحديد",
                "حدّد نتيجة واحدة على الأقل (نقر مزدوج على الصف، أو زر «تحديد الكل»).",
            )
            return
        try:
            import docx  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "مكتبة ناقصة",
                "النسخ إلى Word يحتاج مكتبة python-docx.\nافتح موجّه الأوامر ونفّذ:\n\n"
                "pip install python-docx",
            )
            return

        default_name = (self.q_var.get().strip()[:30] or "ناسخ") + ".docx"
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("مستند Word", "*.docx")],
            title="حفظ ملف Word",
        )
        if not path:
            return
        self._set_status("جارٍ إنشاء ملف Word…")
        threading.Thread(target=self._word_worker, args=(path,), daemon=True).start()

    def _word_worker(self, path: str) -> None:
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        def set_rtl(paragraph):
            pPr = paragraph._p.get_or_add_pPr()
            bidi = OxmlElement("w:bidi")
            pPr.append(bidi)

        doc = Document()
        # عنوان الوثيقة
        q = self.q_var.get().strip()
        title_p = doc.add_paragraph()
        set_rtl(title_p)
        run = title_p.add_run(f"نتائج البحث عن: «{q}»")
        run.bold = True
        run.font.size = Pt(14)
        doc.add_paragraph("")

        for idx, hit in enumerate(self.checked.values(), 1):
            text = aljam3.clean_text(hit.content)
            cit = aljam3.format_citation(hit, with_author=self.with_author_var.get())

            # النص
            body = doc.add_paragraph()
            body.alignment = 3  # justify
            set_rtl(body)
            body.add_run(f"{idx}) ").bold = True
            body.add_run(text)

            # التخريج
            ref = doc.add_paragraph()
            set_rtl(ref)
            r = ref.add_run(f"المصدر: ({cit})")
            r.italic = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

            # الرابط
            if self.with_link_var.get() and hit.book:
                link_p = doc.add_paragraph()
                set_rtl(link_p)
                lr = link_p.add_run(hit.site_url)
                lr.font.size = Pt(9)
                lr.font.color.rgb = RGBColor(0x1A, 0x5F, 0xB4)

            doc.add_paragraph("")

        try:
            doc.save(path)
        except Exception as e:  # noqa: BLE001
            self.after(0, lambda: messagebox.showerror("خطأ", f"تعذّر الحفظ:\n{e}"))
            return

        def done():
            self._set_status(f"تم إنشاء: {path}")
            if messagebox.askyesno("تم بنجاح ✅", f"أُنشئ الملف:\n{path}\n\nهل تريد فتحه الآن؟"):
                try:
                    os.startfile(path)  # ويندوز
                except AttributeError:
                    import subprocess, sys
                    opener = "open" if sys.platform == "darwin" else "xdg-open"
                    subprocess.Popen([opener, path])

        self.after(0, done)


if __name__ == "__main__":
    TakhrijApp().mainloop()
