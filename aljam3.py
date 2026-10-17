#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aljam3.py — عميل للتعامل مع "الجامع" (aljam3.com)
=================================================

الجامع: مكتبة المكتبات الإسلامية (أكثر من 63,500 كتاب). مفتوح ولا يحتاج مفتاحًا.

يستخدم هذا العميل خادم MCP العام للجامع لأنه يُرجع في نتيجة البحث الواحدة:
المقتطف + اسم الكتاب + اسم المؤلف + القسم + المكتبة + رقم الصفحة + روابط
صفحة الجامع وملف الـ PDF — أي كل ما يلزم للتخريج دفعة واحدة.

    نقطة الاتصال:  https://aljam3.com/mcp   (Streamable HTTP، بلا مصادقة)
    البروتوكول:    JSON-RPC 2.0  (tools/call)
    الأدوات:       search, get_page, list_categories, ... (سبع أدوات)

ملاحظات توثيقية (ذكرها الموقع نفسه):
  - رقم الصفحة يخصّ الملف الرقمي وقد يختلف عن ترقيم المطبوع.
  - الجامع لا يوفّر الناشر/الطبعة/درجة الحديث؛ لا تُلفّق هذه البيانات.
  - قد يحتوي النص أخطاء استخراج/OCR؛ راجع رابط الـ PDF عند التوثيق المهم.
"""

from __future__ import annotations

import html as html_lib
import itertools
import json
import re
from dataclasses import dataclass
from functools import lru_cache
from typing import Any, Optional

import requests

# قاعدة REST (نستعملها لجلب اسم المؤلف من بيانات الكتاب عند غيابه في نتيجة البحث)
REST_BASE = "https://aljam3.com/api/v1"

MCP_URL = "https://aljam3.com/mcp"
SITE_BASE = "https://aljam3.com"

_HEADERS = {
    # خادم MCP يتطلب هذين النوعين في Accept
    "Accept": "application/json, text/event-stream",
    "Content-Type": "application/json",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36"
    ),
}

_id_counter = itertools.count(1)


# ---------------------------------------------------------------------------
# بنى البيانات
# ---------------------------------------------------------------------------
@dataclass
class SearchHit:
    """نتيجة بحث: مقتطف مطابق مع بيانات كتابه الكاملة للتخريج."""
    excerpt: str = ""            # المقتطف
    book_name: str = ""
    author_name: str = ""
    category_name: str = ""
    library_name: str = ""
    page_number: Optional[int] = None
    vol: str = ""               # رقم الجزء/المجلد (إن وُجد)
    file_id: Optional[int] = None
    book_id: Optional[int] = None
    url: str = ""               # رابط صفحة الجامع
    pdf_url: str = ""           # رابط صفحة الـ PDF

    @property
    def site_url(self) -> str:
        return self.url or (f"{SITE_BASE}/books/{self.book_id}" if self.book_id else SITE_BASE)


# ---------------------------------------------------------------------------
# طبقة MCP (JSON-RPC)
# ---------------------------------------------------------------------------
def _parse_mcp_body(resp: requests.Response) -> dict[str, Any]:
    """
    يحلّل رد خادم MCP سواء كان JSON صِرفًا أو تدفّق أحداث (SSE: أسطر data:).
    """
    ctype = resp.headers.get("Content-Type", "")
    text = resp.text
    if "text/event-stream" in ctype:
        # اجمع أسطر data: وادمج أول JSON صالح
        for line in text.splitlines():
            line = line.strip()
            if line.startswith("data:"):
                payload = line[len("data:"):].strip()
                if payload:
                    try:
                        return json.loads(payload)
                    except json.JSONDecodeError:
                        continue
        raise ValueError("لم يُعثر على بيانات JSON في تدفّق الأحداث.")
    return resp.json()


def _mcp_call(tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """ينفّذ أداة MCP ويُعيد الكائن structuredContent (أو محتوى النص)."""
    body = {
        "jsonrpc": "2.0",
        "id": next(_id_counter),
        "method": "tools/call",
        "params": {"name": tool, "arguments": arguments},
    }
    resp = requests.post(MCP_URL, headers=_HEADERS, json=body, timeout=30)
    resp.raise_for_status()
    data = _parse_mcp_body(resp)

    if "error" in data and data["error"]:
        err = data["error"]
        raise RuntimeError(f"خطأ من الجامع: {err.get('message', err)}")

    result = data.get("result", {}) or {}
    structured = result.get("structuredContent")
    if structured is not None:
        return structured

    # احتياط: بعض الردود تضع النتيجة كنص JSON داخل content
    for block in result.get("content", []) or []:
        if block.get("type") == "text":
            try:
                return json.loads(block.get("text", "{}"))
            except json.JSONDecodeError:
                pass
    return result


# ---------------------------------------------------------------------------
# دوال الواجهة
# ---------------------------------------------------------------------------
def _first(d: dict, *keys, default=""):
    """يُعيد أول قيمة غير فارغة من بين المفاتيح المحتملة."""
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return default


def _name_of(value: Any) -> str:
    """
    يستخرج الاسم سواء كانت القيمة نصًّا مباشرًا أو كائنًا فيه {'name': ...}.
    (الجامع يُرجع author/category/library ككائنات فيها name و id.)
    """
    if isinstance(value, dict):
        return str(value.get("name", "") or value.get("title", "") or "")
    if value in (None, ""):
        return ""
    return str(value)


def _id_of(value: Any) -> Optional[int]:
    """يستخرج المعرّف من كائن {'id': ...} أو من رقم مباشر."""
    if isinstance(value, dict):
        v = value.get("id")
        return int(v) if isinstance(v, int) else None
    if isinstance(value, int):
        return value
    return None


def search(
    query: str,
    *,
    book_id: Optional[int] = None,
    author_id: Optional[int] = None,
    category_id: Optional[int] = None,
    library_id: Optional[int] = None,
    per_page: int = 20,
    page: int = 1,
    enrich_author: bool = True,
) -> tuple[dict[str, Any], list[SearchHit]]:
    """
    يبحث في نصوص كتب الجامع عبر أداة MCP «search».
    يُرجع (معلومات الترقيم, قائمة النتائج) — كل نتيجة تحوي بيانات كتابها كاملة.

    enrich_author=True: يملأ اسم المؤلف الناقص من بيانات الكتاب (طلبات مُخزّنة مؤقتًا).
    """
    args: dict[str, Any] = {"q": query, "per_page": per_page, "page": page}
    if book_id is not None:
        args["book_id"] = book_id
    if author_id is not None:
        args["author_id"] = author_id
    if category_id is not None:
        args["category_id"] = category_id
    if library_id is not None:
        args["library_id"] = library_id

    data = _mcp_call("search", args)
    pagination = data.get("pagination", {}) or {}
    hits: list[SearchHit] = []

    # قائمة النتائج قد تأتي تحت أحد هذه المفاتيح
    rows = data.get("results") or data.get("pages") or data.get("books") or []
    for r in rows:
        book_id = _id_of(r.get("book")) or _first(r, "book_id", "id", default=None)
        url = _first(r, "url", "aljam3_url", "link")
        file_id = _first(r, "file_id", default=None)
        page_number = _first(r, "page_number", "number", "page", default=None)

        # احتياط: استخرج file_id ورقم الصفحة من الرابط إن لزم
        # نمط الرابط: https://aljam3.com/ar/{book}/{file}/{page}
        if (file_id is None or page_number is None) and url:
            m = re.search(r"/(\d+)/(\d+)/(\d+)/?$", url)
            if m:
                book_id = book_id or int(m.group(1))
                file_id = file_id if file_id is not None else int(m.group(2))
                page_number = page_number if page_number is not None else int(m.group(3))

        hits.append(
            SearchHit(
                excerpt=_first(r, "excerpt", "snip", "text", "content"),
                book_name=_name_of(_first(r, "book", "book_name", "title", "book_title")),
                author_name=_name_of(_first(r, "author", "author_name")),
                category_name=_name_of(_first(r, "category", "category_name", "subject")),
                library_name=_name_of(_first(r, "library", "library_name")),
                page_number=page_number,
                vol=str(_first(r, "vol", "volume", "part", default="")),
                file_id=file_id,
                book_id=book_id,
                url=url,
                pdf_url=_first(r, "pdf_url", "pdf_link", "pdf"),
            )
        )

    if enrich_author:
        enrich_authors(hits)

    return pagination, hits


@lru_cache(maxsize=1024)
def author_of_book(book_id: int) -> str:
    """
    يجلب اسم مؤلف كتاب من واجهة REST (/books/{id}) عند غيابه في نتيجة البحث.
    مُخزَّن مؤقتًا لتفادي الطلبات المكررة لنفس الكتاب.
    """
    try:
        resp = requests.get(
            f"{REST_BASE}/books/{book_id}",
            headers={"Accept": "application/json", "User-Agent": _HEADERS["User-Agent"]},
            timeout=20,
        )
        resp.raise_for_status()
        data = resp.json()
        return _name_of(data.get("author"))
    except requests.RequestException:
        return ""


def enrich_authors(hits: list["SearchHit"]) -> None:
    """يملأ أسماء المؤلفين الناقصة في النتائج (في مكانها)."""
    for hit in hits:
        if not clean_text(hit.author_name) and hit.book_id:
            hit.author_name = author_of_book(int(hit.book_id))


def get_page(file_id: int, page_number: int, max_chars: int = 12000) -> dict[str, Any]:
    """يقرأ نص صفحة بعينها كاملًا عبر أداة MCP «get_page» (يُرجع الكائن الخام)."""
    return _mcp_call(
        "get_page",
        {"file_id": file_id, "page_number": page_number, "max_chars": max_chars},
    )


def get_page_text(file_id: int, page_number: int, max_chars: int = 12000) -> str:
    """
    يقرأ نص صفحة بعينها ويُعيد النص النظيف فقط (نصًّا عاديًا بلا وسوم ولا JSON).
    يستخرج حقل النص من رد get_page أيًّا كان اسمه (text/content).
    """
    data = get_page(file_id, page_number, max_chars=max_chars)
    raw = ""
    if isinstance(data, dict):
        # النص قد يكون في text أو content، وقد يكون داخل كائن page
        raw = _first(data, "text", "content", default="")
        if not raw and isinstance(data.get("page"), dict):
            raw = _first(data["page"], "text", "content", default="")
    else:
        raw = str(data)
    return clean_text(raw)


def list_categories(query: str = "", per_page: int = 20) -> list[dict[str, Any]]:
    """يجلب/يبحث الأقسام عبر أداة MCP «list_categories»."""
    args: dict[str, Any] = {"per_page": per_page}
    if query:
        args["q"] = query
    data = _mcp_call("list_categories", args)
    return data.get("categories", []) or []


def list_libraries(per_page: int = 20) -> list[dict[str, Any]]:
    """يجلب المكتبات المصدرية عبر أداة MCP «list_libraries»."""
    data = _mcp_call("list_libraries", {"per_page": per_page})
    return data.get("libraries", []) or []


# ---------------------------------------------------------------------------
# أدوات النص والتخريج
# ---------------------------------------------------------------------------
_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"[ \t]+")
# مُعلِّمات أرقام الصفحات الخاصة بالجامع مثل: ⦗...⦘  أو  (رقم)  في بداية السطر
_PAGE_MARKER_RE = re.compile(r"⦗[^⦘]*⦘")


def clean_text(raw: str) -> str:
    """
    ينظّف نص الجامع تنظيفًا كاملًا:
      - يزيل وسوم HTML (بما فيها <mark>)
      - يفكّ كيانات HTML (&amp; ...)
      - يحوّل \\r\\n الحرفية (المكتوبة كنص) إلى أسطر فعلية
      - يزيل معلِّمات أرقام الصفحات ⦗...⦘
      - يضبط الفراغات والأسطر
    """
    if not raw:
        return ""
    txt = str(raw)
    # حوّل تسلسلات \r\n و \n و \t المكتوبة حرفيًا (نصًّا) إلى ما يقابلها
    txt = txt.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\\r", "\n")
    txt = txt.replace("\\t", " ")
    txt = _TAG_RE.sub("", txt)
    txt = html_lib.unescape(txt)
    txt = _PAGE_MARKER_RE.sub("", txt)
    # اضبط الفراغات داخل كل سطر، وقلّل الأسطر الفارغة المتتالية
    lines = [_WS_RE.sub(" ", ln).strip() for ln in txt.split("\n")]
    out: list[str] = []
    for ln in lines:
        if ln or (out and out[-1]):   # لا تكرّر الأسطر الفارغة
            out.append(ln)
    return "\n".join(out).strip()


def format_citation(hit: SearchHit, *, with_author: bool = True) -> str:
    """
    يبني المرجع بالصيغة المطلوبة:
        <اسم الكتاب الجزء/الصفحة>
    وإن لم يوجد جزء يُكتفى برقم الصفحة:
        <اسم الكتاب الصفحة>

    (المعامل with_author مُبقًى للتوافق لكنه غير مستخدم في هذه الصيغة.)
    """
    name = clean_text(hit.book_name) or "كتاب غير معروف"
    vol = str(hit.vol).strip() if hit.vol else ""
    page = hit.page_number

    if vol and vol not in ("0", "None"):
        loc = f"{vol}/{page}" if page is not None else vol
    else:
        loc = str(page) if page is not None else ""

    return f"<{name} {loc}>".replace("  ", " ").strip() if loc else f"<{name}>"


# ---------------------------------------------------------------------------
# تشغيل تجريبي
# ---------------------------------------------------------------------------
def _demo(query: str, limit: int = 5) -> None:
    print(f"🔎 البحث في الجامع عن: «{query}»\n" + "=" * 60)
    try:
        pagination, hits = search(query, per_page=limit)
    except Exception as e:  # noqa: BLE001
        print(f"❌ تعذّر الاتصال بالجامع: {e}")
        return
    print(f"إجمالي النتائج: {pagination.get('count', pagination.get('total', '?'))}\n")
    for i, hit in enumerate(hits, 1):
        print(f"[{i}] المرجع: {format_citation(hit)}")
        print(f"    {clean_text(hit.excerpt)[:150]}…")
        print("-" * 60)


def _diagnose(query: str) -> None:
    """يطبع المفاتيح الخام لأول نتيجة (للتشخيص عند الحاجة)."""
    args = {"q": query, "per_page": 1, "page": 1}
    data = _mcp_call("search", args)
    rows = data.get("results") or data.get("pages") or data.get("books") or []
    print("مفاتيح الترقيم:", list((data.get("pagination") or {}).keys()))
    if rows:
        print("أول نتيجة (خام):")
        print(json.dumps(rows[0], ensure_ascii=False, indent=2))
    else:
        print("لا نتائج. المفاتيح العليا:", list(data.keys()))


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--diagnose":
        _diagnose(sys.argv[2] if len(sys.argv) > 2 else "إنما الأعمال بالنيات")
    else:
        q = sys.argv[1] if len(sys.argv) > 1 else "إنما الأعمال بالنيات"
        _demo(q)
