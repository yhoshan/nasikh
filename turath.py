#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
turath.py — عميل بسيط للتعامل مع واجهة turath.io البرمجية (API)
=================================================================

يوفّر هذا الملف دوالّ للبحث في مكتبة التراث (بديل المكتبة الشاملة) واسترجاع
الصفحات مع بياناتها الوصفية (اسم الكتاب، المؤلف، الجزء، الصفحة)، ثم تنسيق
"حاشية تخريج" جاهزة للإدراج في Word بالشكل:

    (الكتاب: اسم الكتاب — ج X / ص Y)

البنية مبنية على واجهة turath.io العامة:
    - القاعدة:      https://api.turath.io/
    - البحث:        /search?q=...&ver=3
    - الصفحة:       /page?book_id=...&pg=...&ver=3

لا يحتاج مفتاح API. يحتاج فقط اتصالاً بالإنترنت وحزمة requests.
"""

from __future__ import annotations

import json
import re
import html as html_lib
from dataclasses import dataclass, field
from typing import Any, Optional

import requests

API_BASE = "https://api.turath.io"
API_VERSION = 3

# ترويسات تحاكي المتصفح لتفادي الحجب
_HEADERS = {
    "Accept": "application/json",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0 Safari/537.36"
    ),
}


# ---------------------------------------------------------------------------
# بنى البيانات
# ---------------------------------------------------------------------------
@dataclass
class PageMeta:
    """البيانات الوصفية لصفحة: اسم الكتاب والمؤلف والجزء والصفحة."""
    book_name: str = ""
    author_name: str = ""
    vol: str = ""           # الجزء / المجلد
    page: Optional[int] = None   # رقم الصفحة (موافق للمطبوع غالبًا)
    page_id: Optional[int] = None
    headings: list[str] = field(default_factory=list)

    @classmethod
    def from_raw(cls, raw: Any) -> "PageMeta":
        """يبني الكائن من حقل meta (قد يكون نصًّا JSON أو قاموسًا)."""
        if isinstance(raw, str):
            try:
                raw = json.loads(raw)
            except (json.JSONDecodeError, TypeError):
                raw = {}
        if not isinstance(raw, dict):
            raw = {}
        return cls(
            book_name=raw.get("book_name", "") or "",
            author_name=raw.get("author_name", "") or "",
            vol=str(raw.get("vol", "") or ""),
            page=raw.get("page"),
            page_id=raw.get("page_id"),
            headings=raw.get("headings") or [],
        )


@dataclass
class SearchHit:
    """نتيجة بحث واحدة."""
    book_id: int
    author_id: int
    cat_id: int
    snippet: str          # المقتطف المُظلَّل (قد يحوي وسوم <span>)
    text: str             # النص الخام
    meta: PageMeta


@dataclass
class PageResult:
    """صفحة كاملة مع نصها وبياناتها الوصفية."""
    text: str
    meta: PageMeta


# ---------------------------------------------------------------------------
# طبقة الشبكة
# ---------------------------------------------------------------------------
def _get(path: str, params: dict[str, Any]) -> dict[str, Any]:
    params = {**params, "ver": API_VERSION}
    url = f"{API_BASE}{path}"
    resp = requests.get(url, params=params, headers=_HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# دوال الواجهة
# ---------------------------------------------------------------------------
def search(
    query: str,
    *,
    book: Optional[int] = None,
    author: Optional[int] = None,
    category: Optional[int] = None,
    page: Optional[int] = None,
    precision: Optional[int] = None,
) -> tuple[int, list[SearchHit]]:
    """
    يبحث عن عبارة في مكتبة التراث.

    يُرجع (العدد الإجمالي للنتائج, قائمة النتائج).
    """
    params: dict[str, Any] = {"q": query}
    if book is not None:
        params["book"] = book
    if author is not None:
        params["author"] = author
    if category is not None:
        params["cat_id"] = category
    if page is not None:
        params["page"] = page
    if precision is not None:
        params["precision"] = precision

    data = _get("/search", params)
    count = int(data.get("count", 0) or 0)
    hits: list[SearchHit] = []
    for row in data.get("data", []) or []:
        hits.append(
            SearchHit(
                book_id=row.get("book_id", 0),
                author_id=row.get("author_id", 0),
                cat_id=row.get("cat_id", 0),
                snippet=row.get("snip", "") or "",
                text=row.get("text", "") or "",
                meta=PageMeta.from_raw(row.get("meta")),
            )
        )
    return count, hits


def get_page(book_id: int, page_number: int) -> PageResult:
    """يسترجع صفحة محددة من كتاب بنصها وبياناتها الوصفية."""
    data = _get("/page", {"book_id": book_id, "pg": page_number})
    meta_raw = data.get("meta")
    text = data.get("text", "") or ""
    if not meta_raw and not text:
        raise ValueError(f"الصفحة غير موجودة (كتاب {book_id}، صفحة {page_number})")
    return PageResult(text=text, meta=PageMeta.from_raw(meta_raw))


def get_author(author_id: int) -> str:
    """يسترجع نبذة المؤلف (نص)."""
    data = _get("/author", {"id": author_id})
    return data.get("info", "") or ""


def get_book_info(book_id: int) -> dict[str, Any]:
    """
    يسترجع معلومات الكتاب وفهارسه (العناوين، المجلدات، حدود الأجزاء...).
    يُرجع القاموس كما تُعيده الـ API: {'meta': {...}, 'indexes': {...}}.
    """
    return _get("/book", {"id": book_id, "include": "indexes"})


# ---------------------------------------------------------------------------
# أدوات تنظيف النص وتنسيق الحاشية
# ---------------------------------------------------------------------------
_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"[ \t]+")


def clean_text(raw: str) -> str:
    """يزيل وسوم HTML ويفكّ ترميز الكيانات ويضبط الفراغات."""
    if not raw:
        return ""
    txt = _TAG_RE.sub("", raw)
    txt = html_lib.unescape(txt)
    txt = _WS_RE.sub(" ", txt)
    return txt.strip()


def format_citation(meta: PageMeta, *, with_author: bool = True) -> str:
    """
    يبني نص الحاشية (التخريج) بالشكل:
        اسم الكتاب (لفلان) — ج X / ص Y

    يتجاوز الجزء إن لم يكن موجودًا.
    """
    parts: list[str] = []
    name = meta.book_name or "كتاب غير معروف"
    if with_author and meta.author_name:
        parts.append(f"{name}، لـ{meta.author_name}")
    else:
        parts.append(name)

    loc: list[str] = []
    if meta.vol and str(meta.vol).strip() not in ("", "0"):
        loc.append(f"ج {meta.vol}")
    if meta.page is not None:
        loc.append(f"ص {meta.page}")
    if loc:
        parts.append(" / ".join(loc))

    return " — ".join(parts)


# ---------------------------------------------------------------------------
# تشغيل تجريبي من سطر الأوامر
# ---------------------------------------------------------------------------
def _demo(query: str, limit: int = 5) -> None:
    print(f"🔎 البحث عن: «{query}»\n" + "=" * 60)
    try:
        count, hits = search(query)
    except requests.RequestException as e:
        print(f"❌ تعذّر الاتصال بالـ API: {e}")
        print("   (تأكد من اتصال الإنترنت ومن أن النطاق api.turath.io غير محجوب)")
        return

    print(f"إجمالي النتائج: {count}\n")
    for i, hit in enumerate(hits[:limit], 1):
        citation = format_citation(hit.meta)
        snippet = clean_text(hit.snippet) or clean_text(hit.text)
        print(f"[{i}] {citation}")
        print(f"    المقتطف: {snippet[:160]}{'…' if len(snippet) > 160 else ''}")
        print(f"    الحاشية الجاهزة:  ({citation})")
        print("-" * 60)


if __name__ == "__main__":
    import sys

    q = sys.argv[1] if len(sys.argv) > 1 else "إنما الأعمال بالنيات"
    _demo(q)
