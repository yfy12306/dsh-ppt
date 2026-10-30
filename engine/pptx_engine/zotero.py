# -*- coding: utf-8 -*-
"""Zotero 集成 —— 抽取自上游 cli_anything/zotero，为「做 PPT」这一目标收窄。

上游那套有 8 个 core 模块 + 38KB CLI，覆盖文献管理全流程。本模块只保留
**做汇报 PPT 真正用得到的三件事**：

1. **找文献**：直读 `zotero.sqlite`（只读连接，绝不写库），按标题/作者/年份/标签检索。
2. **出引用**：本地实现 APA / GB/T 7714 / IEEE 三种格式。
   —— 这里刻意**不依赖** Zotero Local API 的 citation 端点：上游
   `item citation` 要求 Zotero 正在运行且开了 Local API，而做 PPT 时
   Zotero 往往没开。本地渲染只要 sqlite 文件在就能用。
3. **生成参考文献页**：把检索结果直接排成一页 slide spec，交给 renderer。

Zotero 未安装时全部函数优雅降级（返回 available=False），不影响 PPT 主流程。

数据目录探测沿用上游 zotero_paths.py 的思路：
    profiles.ini → 活动 profile → prefs.js 的 extensions.zotero.dataDir
    → 回退 ~/Zotero
"""
from __future__ import annotations

import configparser
import json
import os
import re
import sqlite3
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Optional

DATA_DIR_PREF = "extensions.zotero.dataDir"
USE_DATA_DIR_PREF = "extensions.zotero.useDataDir"
LOCAL_API_PREF = "extensions.zotero.httpServer.localAPI.enabled"
DEFAULT_HTTP_PORT = 23119


def _log(msg: str) -> None:
    import sys
    print(f"[zotero] {msg}", file=sys.stderr, flush=True)


# ---------------------------------------------------------------- 环境探测
def candidate_profile_roots() -> list[Path]:
    home = Path.home()
    out = [
        home / "AppData" / "Roaming" / "Zotero" / "Zotero" / "Profiles",   # Windows
        home / ".zotero" / "integration",                                    # Linux
        home / "Library" / "Application Support" / "Zotero" / "Profiles",    # macOS
    ]
    env = os.environ.get("ZOTERO_PROFILE_ROOT")
    if env:
        out.insert(0, Path(env))
    return out


def find_profile_root() -> Optional[Path]:
    for p in candidate_profile_roots():
        if p.is_dir() and (p / "profiles.ini").exists():
            return p
    return None


def find_active_profile(root: Optional[Path] = None) -> Optional[Path]:
    root = root or find_profile_root()
    if root is None:
        return None
    cp = configparser.ConfigParser(strict=False)
    try:
        cp.read(str(root / "profiles.ini"), encoding="utf-8")
    except Exception:  # noqa: BLE001
        return None
    # 优先 default= 指向的 section
    general = cp["General"] if cp.has_section("General") else {}
    default_name = str(general.get("default", "")).strip()
    sections = [s for s in cp.sections() if s.startswith("Profile")]
    ordered = ([f"Profile{default_name}"] if default_name else []) + \
              [s for s in sections if s != f"Profile{default_name}"]
    for sec in ordered:
        if not cp.has_section(sec):
            continue
        rel = cp[sec].get("Path", "").strip()
        is_rel = cp[sec].get("IsRelative", "0").strip() == "1"
        if not rel:
            continue
        p = (root / rel) if is_rel else Path(rel)
        if (p / "prefs.js").exists() or p.is_dir():
            return p
    return None


def _read_pref(profile_dir: Optional[Path], name: str) -> Optional[str]:
    if profile_dir is None:
        return None
    for fn in ("user.js", "prefs.js"):
        f = profile_dir / fn
        if not f.exists():
            continue
        try:
            txt = f.read_text(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            continue
        m = re.search(r'user_pref\(\s*"' + re.escape(name) + r'"\s*,\s*(.+?)\s*\)\s*;', txt)
        if not m:
            continue
        raw = m.group(1).strip()
        if raw.startswith('"') and raw.endswith('"'):
            try:
                return json.loads(raw)
            except Exception:  # noqa: BLE001
                return raw[1:-1]
        return raw
    return None


def find_data_dir(profile_dir: Optional[Path] = None) -> Optional[Path]:
    """定位 Zotero 数据目录（含 zotero.sqlite）。"""
    env = os.environ.get("ZOTERO_DATA_DIR")
    if env and Path(env).is_dir():
        return Path(env)
    pd = profile_dir if profile_dir is not None else find_active_profile()
    use = _read_pref(pd, USE_DATA_DIR_PREF)
    if str(use).lower() in ("true", "1"):
        d = _read_pref(pd, DATA_DIR_PREF)
        if d:
            p = Path(d)
            # prefs.js 里的路径可能是转义过的 Windows 路径
            p = Path(str(p).replace("\\\\", "\\"))
            if p.is_dir():
                return p
    home = Path.home()
    for c in (home / "Zotero", home / "Documents" / "Zotero",
              home / "AppData" / "Local" / "Zotero"):
        if c.is_dir() and (c / "zotero.sqlite").exists():
            return c
    return None


def http_port(profile_dir: Optional[Path] = None) -> int:
    pd = profile_dir if profile_dir is not None else find_active_profile()
    v = _read_pref(pd, "extensions.zotero.httpServer.port")
    try:
        return int(v) if v else DEFAULT_HTTP_PORT
    except (TypeError, ValueError):
        return DEFAULT_HTTP_PORT


def local_api_enabled(profile_dir: Optional[Path] = None) -> bool:
    pd = profile_dir if profile_dir is not None else find_active_profile()
    return str(_read_pref(pd, LOCAL_API_PREF)).lower() in ("true", "1")


def probe() -> dict:
    """一次性探测 Zotero 环境，供 zotero_status 工具返回。"""
    root = find_profile_root()
    prof = find_active_profile(root)
    data = find_data_dir(prof)
    sqlite_path = (data / "zotero.sqlite") if data else None
    port = http_port(prof)
    api_on = local_api_enabled(prof)

    running = False
    if api_on or True:  # 即便 pref 没开也探一次，Zotero 7 默认可能已开
        try:
            req = urllib.request.Request(f"http://127.0.0.1:{port}/api/users/0/items?limit=1",
                                         headers={"Zotero-API-Version": "3"})
            with urllib.request.urlopen(req, timeout=2) as r:
                running = r.status == 200
        except Exception:  # noqa: BLE001
            running = False

    out = {
        "installed": bool(data and sqlite_path and sqlite_path.exists()),
        "profileRoot": str(root) if root else None,
        "profileDir": str(prof) if prof else None,
        "dataDir": str(data) if data else None,
        "sqlitePath": str(sqlite_path) if sqlite_path else None,
        "sqliteExists": bool(sqlite_path and sqlite_path.exists()),
        "httpPort": port,
        "localApiEnabled": api_on,
        "localApiRunning": running,
        "itemCount": None,
        "error": "",
    }
    if out["sqliteExists"]:
        try:
            with _connect(sqlite_path) as conn:
                row = conn.execute(
                    "SELECT COUNT(*) FROM items WHERE itemTypeID NOT IN (1, 14) "
                    "AND itemID NOT IN (SELECT itemID FROM deletedItems)").fetchone()
                out["itemCount"] = int(row[0]) if row else 0
        except Exception as e:  # noqa: BLE001
            out["error"] = f"读库失败：{type(e).__name__}: {e}"
    else:
        out["error"] = "未找到 zotero.sqlite（Zotero 未安装，或数据目录在自定义位置——设 ZOTERO_DATA_DIR 环境变量）"
    return out


# ---------------------------------------------------------------- 只读连接
def _connect(sqlite_path) -> sqlite3.Connection:
    """只读连接。

    Zotero 运行时会对库加写锁；用 `file:...?mode=ro&immutable=1` 才能在其运行时也读到。
    immutable=1 告诉 SQLite 文件不会变，绕过锁 —— 代价是可能读到略旧的数据，
    对「查文献做 PPT」这个场景完全可接受，且**绝不会破坏用户的库**。
    """
    p = Path(sqlite_path).resolve()
    uri = f"file:{urllib.parse.quote(str(p))}?mode=ro&immutable=1"
    conn = sqlite3.connect(uri, uri=True, timeout=5)
    conn.row_factory = sqlite3.Row
    return conn


def _ctx() -> tuple:
    prof = find_active_profile()
    data = find_data_dir(prof)
    if not data:
        raise RuntimeError("未找到 Zotero 数据目录")
    sq = data / "zotero.sqlite"
    if not sq.exists():
        raise RuntimeError(f"zotero.sqlite 不存在：{sq}")
    return prof, data, sq


# ---------------------------------------------------------------- 查询
_ITEM_BASE = """
SELECT i.itemID, i.key, i.itemTypeID, it.typeName,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'title') AS title,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'date') AS date,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'publicationTitle') AS publication,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'DOI') AS doi,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'url') AS url,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'volume') AS volume,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'issue') AS issue,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'pages') AS pages,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'publisher') AS publisher,
       (SELECT v.value FROM itemData d
          JOIN fields f ON f.fieldID = d.fieldID
          JOIN itemDataValues v ON v.valueID = d.valueID
         WHERE d.itemID = i.itemID AND f.fieldName = 'abstractNote') AS abstract
FROM items i
JOIN itemTypes it ON it.itemTypeID = i.itemTypeID
"""


def _creators(conn, item_id: int) -> list[dict]:
    rows = conn.execute(
        """SELECT c.firstName, c.lastName, c.fieldMode, ic.creatorTypeID, ic.orderIndex
           FROM itemCreators ic
           JOIN creators c ON c.creatorID = ic.creatorID
          WHERE ic.itemID = ? ORDER BY ic.orderIndex""", (item_id,)).fetchall()
    out = []
    for r in rows:
        if int(r["fieldMode"] or 0) == 1:
            name = (r["lastName"] or "").strip()
            first = ""
        else:
            name = (r["lastName"] or "").strip()
            first = (r["firstName"] or "").strip()
        out.append({"last": name, "first": first,
                    "typeId": int(r["creatorTypeID"] or 1),
                    "index": int(r["orderIndex"] or 0)})
    return out


def _tags(conn, item_id: int) -> list[str]:
    rows = conn.execute(
        """SELECT t.name FROM itemTags itg JOIN tags t ON t.tagID = itg.tagID
           WHERE itg.itemID = ?""", (item_id,)).fetchall()
    return [r["name"] for r in rows if r["name"]]


def _collections(conn, item_id: int) -> list[str]:
    rows = conn.execute(
        """SELECT c.collectionName FROM collectionItems ci
           JOIN collections c ON c.collectionID = ci.collectionID
          WHERE ci.itemID = ?""", (item_id,)).fetchall()
    return [r["collectionName"] for r in rows if r["collectionName"]]


def _row_to_item(conn, row) -> dict:
    iid = int(row["itemID"])
    date = row["date"] or ""
    year = ""
    m = re.search(r"(1[89]\d{2}|20\d{2})", str(date))
    if m:
        year = m.group(1)
    return {
        "id": iid,
        "key": row["key"],
        "type": row["typeName"],
        "title": row["title"] or "",
        "date": str(date),
        "year": year,
        "publication": row["publication"] or "",
        "publisher": row["publisher"] or "",
        "volume": row["volume"] or "",
        "issue": row["issue"] or "",
        "pages": row["pages"] or "",
        "doi": row["doi"] or "",
        "url": row["url"] or "",
        "abstract": row["abstract"] or "",
        "creators": _creators(conn, iid),
        "tags": _tags(conn, iid),
        "collections": _collections(conn, iid),
    }


def search(query: str = "", limit: int = 20, year: Optional[str] = None,
           item_type: Optional[str] = None, tag: Optional[str] = None,
           collection: Optional[str] = None) -> dict:
    """检索文献。query 为空则列最近添加的条目。"""
    try:
        _p, _d, sq = _ctx()
    except RuntimeError as e:
        return {"available": False, "error": str(e), "items": [], "count": 0}

    where = ["i.itemTypeID NOT IN (1, 14)",           # 排除 attachment / note
             "i.itemID NOT IN (SELECT itemID FROM deletedItems)"]
    args: list[Any] = []
    if query:
        like = f"%{query}%"
        where.append("""(
            EXISTS (SELECT 1 FROM itemData d JOIN fields f ON f.fieldID=d.fieldID
                     JOIN itemDataValues v ON v.valueID=d.valueID
                    WHERE d.itemID=i.itemID AND f.fieldName IN ('title','publicationTitle')
                      AND v.value LIKE ?)
         OR EXISTS (SELECT 1 FROM itemCreators ic JOIN creators c ON c.creatorID=ic.creatorID
                    WHERE ic.itemID=i.itemID
                      AND (c.lastName LIKE ? OR c.firstName LIKE ?))
         OR EXISTS (SELECT 1 FROM itemTags itg JOIN tags t ON t.tagID=itg.tagID
                    WHERE itg.itemID=i.itemID AND t.name LIKE ?)
        )""")
        args += [like, like, like, like]
    if year:
        where.append("""EXISTS (SELECT 1 FROM itemData d JOIN fields f ON f.fieldID=d.fieldID
                        JOIN itemDataValues v ON v.valueID=d.valueID
                       WHERE d.itemID=i.itemID AND f.fieldName='date' AND v.value LIKE ?)""")
        args.append(f"%{year}%")
    if item_type:
        where.append("it.typeName = ?")
        args.append(item_type)
    if tag:
        where.append("""EXISTS (SELECT 1 FROM itemTags itg JOIN tags t ON t.tagID=itg.tagID
                       WHERE itg.itemID=i.itemID AND t.name = ?)""")
        args.append(tag)
    if collection:
        where.append("""EXISTS (SELECT 1 FROM collectionItems ci JOIN collections c
                         ON c.collectionID=ci.collectionID
                       WHERE ci.itemID=i.itemID AND c.collectionName LIKE ?)""")
        args.append(f"%{collection}%")

    sql = _ITEM_BASE + " WHERE " + " AND ".join(where) + \
        " ORDER BY i.dateAdded DESC LIMIT ?"
    args.append(int(limit))

    try:
        with _connect(sq) as conn:
            rows = conn.execute(sql, args).fetchall()
            items = [_row_to_item(conn, r) for r in rows]
    except Exception as e:  # noqa: BLE001
        return {"available": False, "error": f"{type(e).__name__}: {e}", "items": [], "count": 0}
    return {"available": True, "error": "", "count": len(items), "items": items,
            "query": query, "limit": limit}


def get_item(ref: str) -> dict:
    """按 key 或 itemID 取单条。"""
    try:
        _p, _d, sq = _ctx()
    except RuntimeError as e:
        return {"available": False, "error": str(e), "item": None}
    try:
        with _connect(sq) as conn:
            if str(ref).isdigit():
                row = conn.execute(_ITEM_BASE + " WHERE i.itemID = ?", (int(ref),)).fetchone()
            else:
                row = conn.execute(_ITEM_BASE + " WHERE i.key = ?", (str(ref),)).fetchone()
            if row is None:
                return {"available": True, "error": f"未找到条目：{ref}", "item": None}
            return {"available": True, "error": "", "item": _row_to_item(conn, row)}
    except Exception as e:  # noqa: BLE001
        return {"available": False, "error": f"{type(e).__name__}: {e}", "item": None}


def collections(limit: int = 100) -> dict:
    try:
        _p, _d, sq = _ctx()
        with _connect(sq) as conn:
            rows = conn.execute(
                """SELECT c.collectionID, c.key, c.collectionName, c.parentCollectionID,
                          (SELECT COUNT(*) FROM collectionItems ci WHERE ci.collectionID=c.collectionID) AS n
                   FROM collections c ORDER BY c.collectionName LIMIT ?""", (int(limit),)).fetchall()
            return {"available": True, "error": "",
                    "collections": [dict(r) for r in rows]}
    except Exception as e:  # noqa: BLE001
        return {"available": False, "error": f"{type(e).__name__}: {e}", "collections": []}


# ---------------------------------------------------------------- 引用渲染
def _is_cjk(s: str) -> bool:
    """判断字符串是否含 CJK 字符 —— 中日韩姓名不能按西方规则缩写成首字母。"""
    return any("\u3040" <= ch <= "\u30ff" or "\u4e00" <= ch <= "\u9fff"
               or "\uac00" <= ch <= "\ud7af" for ch in (s or ""))


def _author_apa(creators: list[dict], max_n: int = 6) -> str:
    auth = [c for c in creators if c.get("typeId") == 1] or creators
    if not auth:
        return ""

    def one(c):
        last = (c.get("last") or "").strip()
        first = (c.get("first") or "").strip()
        if not first:
            return last
        # 中文姓名：姓+名直接连写，不缩写、不加逗号（「于飞洋」而非「于, 飞.」）
        if _is_cjk(last) or _is_cjk(first):
            return f"{last}{first}"
        initials = "".join(f"{p[0]}." for p in re.split(r"[\s\-]+", first) if p)
        return f"{last}, {initials}" if initials else last

    if len(auth) == 1:
        return one(auth[0])
    if len(auth) == 2:
        return f"{one(auth[0])}, & {one(auth[1])}"
    if len(auth) <= max_n:
        return ", ".join(one(c) for c in auth[:-1]) + f", & {one(auth[-1])}"
    return ", ".join(one(c) for c in auth[:max_n - 1]) + f", ... {one(auth[-1])}"


def _author_gbt(creators: list[dict], max_n: int = 3) -> str:
    """GB/T 7714：中文姓名连写全名；西文姓名姓前名后、名缩写不加句点；超 3 人用「等」。"""
    auth = [c for c in creators if c.get("typeId") == 1] or creators
    if not auth:
        return ""

    def one(c):
        last = (c.get("last") or "").strip()
        first = (c.get("first") or "").strip()
        if not first:
            return last
        if _is_cjk(last) or _is_cjk(first):
            return f"{last}{first}"          # 于飞洋
        ini = "".join(p[0].upper() for p in re.split(r"[\s\-]+", first) if p)
        return f"{last} {ini}".strip()       # Smith J

    if len(auth) <= max_n:
        return ", ".join(one(c) for c in auth)
    return ", ".join(one(c) for c in auth[:max_n]) + ", 等"


def _author_plain(creators: list[dict]) -> str:
    auth = [c for c in creators if c.get("typeId") == 1] or creators
    out = []
    for c in auth:
        last = (c.get("last") or "").strip()
        first = (c.get("first") or "").strip()
        out.append(f"{last}{first}" if (_is_cjk(last) or _is_cjk(first))
                   else f"{first} {last}".strip())
    return ", ".join(out)


_GBT_TYPE = {
    "journalArticle": "J", "conferencePaper": "C", "book": "M", "bookSection": "M",
    "thesis": "D", "report": "R", "patent": "P", "newspaperArticle": "N",
    "webpage": "EB/OL", "preprint": "J", "document": "Z", "presentation": "Z",
}


def render_citation(item: dict, style: str = "gbt7714") -> str:
    """本地渲染引用串。style: gbt7714 | apa | ieee | plain。"""
    style = str(style or "gbt7714").lower().replace("-", "").replace("_", "")
    title = item.get("title", "")
    year = item.get("year") or (item.get("date", "")[:4] if item.get("date") else "")
    pub = item.get("publication") or item.get("publisher") or ""
    vol, iss, pages = item.get("volume", ""), item.get("issue", ""), item.get("pages", "")
    doi = item.get("doi", "")
    typ = item.get("type", "")

    if style in ("apa", "apa7"):
        a = _author_apa(item.get("creators") or [])
        s = f"{a} ({year or 'n.d.'}). {title}."
        if pub:
            s += f" {pub}"
            if vol:
                s += f", {vol}" + (f"({iss})" if iss else "")
            if pages:
                s += f", {pages}"
            s += "."
        if doi:
            s += f" https://doi.org/{doi}"
        return s.strip()

    if style in ("ieee",):
        auth = [c for c in (item.get("creators") or []) if c.get("typeId") == 1] or (item.get("creators") or [])

        def one(c):
            last = (c.get("last") or "").strip()
            f = (c.get("first") or "").strip()
            if not f:
                return last
            if _is_cjk(last) or _is_cjk(f):
                return f"{last}{f}"
            ini = "".join(f"{p[0]}." for p in re.split(r"[\s\-]+", f) if p)
            return f"{ini} {last}".strip()

        if len(auth) > 3:
            a = one(auth[0]) + " et al."
        else:
            a = ", ".join(one(c) for c in auth)
        s = f'{a}, "{title},"'
        if pub:
            s += f" {pub},"
        if vol:
            s += f" vol. {vol},"
        if iss:
            s += f" no. {iss},"
        if pages:
            s += f" pp. {pages},"
        s += f" {year}."
        if doi:
            s += f" doi: {doi}."
        return s.strip()

    if style == "plain":
        return f"{_author_plain(item.get('creators') or [])} ({year}). {title}. {pub}.".strip()

    # 默认 GB/T 7714-2015（中文汇报/论文最常用）
    a = _author_gbt(item.get("creators") or [])
    mark = _GBT_TYPE.get(typ, "Z")
    s = f"{a}. {title}[{mark}]."
    if pub:
        s += f" {pub},"
    if year:
        s += f" {year}"
    if vol:
        s += f",{vol}"
        if iss:
            s += f"({iss})"
    if pages:
        s += f":{pages}"
    s += "."
    if doi:
        s += f" DOI:{doi}."
    return s.strip()


def render_bibliography(items: list, style: str = "gbt7714", numbered: bool = True) -> list[str]:
    out = []
    for i, it in enumerate(items, start=1):
        c = render_citation(it, style)
        out.append(f"[{i}] {c}" if numbered else c)
    return out


# ---------------------------------------------------------------- 生成参考文献页
def refs_slide_spec(items: list, style: str = "gbt7714", title: str = "参考文献",
                    theme: Optional[dict] = None, per_slide: int = 8,
                    sec: str = "") -> list[dict]:
    """把文献列表排成 slide spec（可直接喂给 Deck.build 的 slides[]）。

    超过 per_slide 条自动分页 —— 参考文献页最容易文字溢出，
    这里按条数硬分页而不是靠字号缩到看不见。
    """
    from .themes import get_theme
    th = theme or get_theme()
    primary = th.get("primary", "#1A3C8B")
    cites = render_bibliography(items, style, numbered=True)
    pages = [cites[i:i + per_slide] for i in range(0, len(cites), per_slide)] or [[]]
    slides = []
    for pi, chunk in enumerate(pages, start=1):
        t = title if len(pages) == 1 else f"{title}（{pi}/{len(pages)}）"
        elements = [{"type": "header", "sec": sec, "title": t, "color": primary}]
        y = 84.0
        fs = 13 if per_slide <= 8 else 11.5
        for c in chunk:
            # 按字符数估算行数（中文按 1 字宽、英文按 0.55 字宽粗估）
            width_units = sum(1.0 if ord(ch) > 0x2E80 else 0.55 for ch in c)
            lines = max(1, int(width_units / 78) + 1)
            hgt = lines * (fs * 1.5) + 6
            elements.append({"type": "text", "box": [40, y, 880, hgt], "text": c,
                             "fs": fs, "color": "#222222", "align": "left", "anchor": "top"})
            y += hgt + 4
            if y > 496:
                break
        slides.append({"elements": elements})
    return slides
