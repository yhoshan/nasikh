#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aljam3.py — عميل للتعامل مع واجهة "الجامع" البرمجية الرسمية (REST API v1)
=========================================================================

الجامع (aljam3.com): مكتبة المكتبات الإسلامية (أكثر من 63,500 كتاب).
الواجهة رسمية وموثّقة (Swagger) ولا تحتاج مفتاحًا.

القاعدة:   https://aljam3.com/api/v1
أهم المسارات المستخدمة في برنامج التخريج:
    GET /search?q=...                      → بحث في نصوص كل الكتب
    GET /books/{id}                        → معلومات كتاب (اسم، مؤلف، قسم، مكتبة)
    GET /books/{id}?expand[]=files         → + ملفات الكتاب (المجلدات)
    GET /authors/{id}                      → معلومات مؤلف

ملاحظات توثيقية (ذكرها الموقع نفسه):
  - رقم الصفحة يخصّ الملف الرقمي وقد يختلف عن ترقيم المطبوع.
  - الجامع لا يوفّر الناشر/الطبعة/درجة الحديث؛ لا نُلفّقها.
"""

from __future__ import annotations

import html as html_lib
import re
from dataclasses import dataclass, field
from functools import lru_cache
from typing import Any, Optional

import requests

API_BASE = "https://aljam3.com/api/v1"
SITE_BASE = "https://aljam3.com"

_HEADERS = {
    "Accept": "application/json",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36"
    ),
}


# ---------------------------------------------------------------------------
# بنى البيانات
# ---------------------------------------------------------------------------
@dataclass
class BookInfo:
    id: int
    title: str = ""
    author_name: str = ""
    category_name: str = ""
    library_name: str = ""
    volumes: Optional[int] = None


@dataclass
class SearchHit:
    """نتيجة بحث: صفحة مطابقة مع نصّها ورقمها وبيانات كتابها."""
    page_id: int
    file_id: int
    page_number: int
    content: str                 # نص الصفحة (قد يحوي وسوم <mark> للتظليل)
    book: Optional[BookInfo] = None

    @property
    def site_url(self) -> str:
        """رابط صفحة الجامع لهذه النتيجة (للمراجعة والتوثيق)."""
        return f"{SITE_BASE}/books/{self.book.id}" if self.book else SITE_BASE


# ---------------------------------------------------------------------------
# طبقة الشبكة
# ---------------------------------------------------------------------------
def _get(path: str, params: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    url = f"{API_BASE}{path}"
    resp = requests.get(url, params=params or {}, headers=_HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# دوال الواجهة
# ---------------------------------------------------------------------------
@lru_cache(maxsize=512)
def get_book_info(book_id: int) -> BookInfo:
    """يجلب بيانات كتاب (مع التخزين المؤقت لتفادي الطلبات المكررة)."""
    data = _get(f"/books/{book_id}")
    author = data.get("author") or {}
    category = data.get("category") or {}
    library = data.get("library") or {}
    return BookInfo(
        id=data.get("id", book_id),
        title=clean_text(data.get("title", "")),
        author_name=clean_text(author.get("name", "")),
        category_name=clean_text(category.get("name", "")),
        library_name=clean_text(library.get("name", "")),
        volumes=data.get("volumes") if data.get("volumes", -1) not in (-1, 0) else None,
    )


def _book_id_from_file(file_id: int) -> Optional[int]:
    """يحاول معرفة رقم الكتاب من رقم الملف عبر مسار /files/{id}."""
    try:
        data = _get(f"/files/{file_id}")
        # الملف قد يحمل رابط الكتاب؛ نحاول استخراج المعرّف من الرابط
        link = data.get("book_link") or ""
        m = re.search(r"/books/(\d+)", link)
        if m:
            return int(m.group(1))
    except requests.RequestException:
        pass
    return None


def search(
    query: str,
    *,
    library: Optional[int] = None,
    books: Optional[list[int]] = None,
    authors: Optional[list[int]] = None,
    categories: Optional[list[int]] = None,
    limit: int = 20,
    page: int = 1,
    enrich: bool = True,
) -> tuple[dict[str, Any], list[SearchHit]]:
    """
    يبحث عن عبارة في نصوص كل كتب الجامع.

    enrich=True: يجلب اسم الكتاب والمؤلف لكل نتيجة (طلبات إضافية، لكن مُخزّنة مؤقتًا).
    يُرجع (معلومات الترقيم, قائمة النتائج).
    """
    params: dict[str, Any] = {"q": query, "limit": limit, "page": page}
    if library is not None:
        params["library"] = library
    if books:
        params["books[]"] = books
    if authors:
        params["authors[]"] = authors
    if categories:
        params["categories[]"] = categories

    data = _get("/search", params)
    pagination = data.get("pagination", {}) or {}
    hits: list[SearchHit] = []

    for row in data.get("pages", []) or []:
        hit = SearchHit(
            page_id=row.get("id", 0),
            file_id=row.get("file_id", 0),
            page_number=row.get("number", 0),
            content=row.get("content", "") or "",
        )
        hits.append(hit)

    if enrich:
        for hit in hits:
            bid = _book_id_from_file(hit.file_id)
            if bid:
                try:
                    hit.book = get_book_info(bid)
                except requests.RequestException:
                    hit.book = None

    return pagination, hits


def list_categories() -> list[dict[str, Any]]:
    """يجلب قائمة الأقسام (للفلترة في الواجهة)."""
    data = _get("/categories")
    return data.get("categories", []) or []


def list_libraries() -> list[dict[str, Any]]:
    """يجلب قائمة المكتبات المصدرية."""
    data = _get("/libraries")
    return data.get("libraries", []) or []


# ---------------------------------------------------------------------------
# أدوات النص والتخريج
# ---------------------------------------------------------------------------
_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"[ \t]+")


def clean_text(raw: str) -> str:
    """يزيل الوسوم (بما فيها <mark>) ويفكّ الكيانات ويضبط الفراغات."""
    if not raw:
        return ""
    txt = _TAG_RE.sub("", raw)
    txt = html_lib.unescape(txt)
    txt = _WS_RE.sub(" ", txt)
    return txt.strip()


def format_citation(hit: SearchHit, *, with_author: bool = True) -> str:
    """
    يبني نص التخريج من نتيجة بحث:
        اسم الكتاب، لـفلان (ص رقم)
    """
    if not hit.book:
        return f"الجامع — ص {hit.page_number}"
    name = hit.book.title or "كتاب غير معروف"
    parts: list[str] = []
    if with_author and hit.book.author_name:
        parts.append(f"{name}، لـ{hit.book.author_name}")
    else:
        parts.append(name)
    parts.append(f"ص {hit.page_number}")
    return "، ".join(parts)


# ---------------------------------------------------------------------------
# تشغيل تجريبي
# ---------------------------------------------------------------------------
def _demo(query: str, limit: int = 5) -> None:
    print(f"🔎 البحث في الجامع عن: «{query}»\n" + "=" * 60)
    try:
        pagination, hits = search(query, limit=limit)
    except requests.RequestException as e:
        print(f"❌ تعذّر الاتصال بالجامع: {e}")
        return
    print(f"إجمالي النتائج: {pagination.get('count', '?')}\n")
    for i, hit in enumerate(hits, 1):
        print(f"[{i}] {format_citation(hit)}")
        print(f"    {clean_text(hit.content)[:150]}…")
        print("-" * 60)


if __name__ == "__main__":
    import sys
    q = sys.argv[1] if len(sys.argv) > 1 else "إنما الأعمال بالنيات"
    _demo(q)
