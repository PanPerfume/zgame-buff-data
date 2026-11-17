#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_buff_from_feishu.py
========================
飞书「ZGame大文本」→ buff-index.json / buff-meta.json → GitHub
一条命令替代原来的「手动下载 → 改名 → 跑脚本 → 上传」四步。

用法
----
# 1) 全自动：从飞书文件夹取最新 Zgame ALT*.xlsx，生成并上传
python sync_buff_from_feishu.py

# 2) 只生成、不上传（先验证数据）
python sync_buff_from_feishu.py --dry-run --output-dir ./out

# 3) 用本地已下载的 Excel（今天就能用，不需要飞书权限）
python sync_buff_from_feishu.py --local "C:\\Users\\admin\\Desktop\\Zgame ALT 1.6.54.853.1.xlsx"

# 4) 自动在本机 Desktop / Downloads 里找最新的 Zgame ALT*.xlsx
python sync_buff_from_feishu.py --auto-local

环境变量（推荐；不要把密钥写进脚本）
----
PowerShell:
  $env:FEISHU_APP_ID="cli_xxx"
  $env:FEISHU_APP_SECRET="xxx"
  $env:GITHUB_TOKEN="github_pat_xxx"

可选：
  FEISHU_FOLDER_TOKEN  默认 fldcnvc54hgGywuCsTssbUTlwYg
  GITHUB_OWNER         默认 PanPerfume
  GITHUB_REPO          默认 zgame-buff-data
  GITHUB_BRANCH        默认 main

前置条件（飞书侧，仅全自动模式需要）
----
1. 飞书开放平台 → 该应用 → 权限管理，开通 drive:drive:readonly（只读足够）
2. 创建并发布应用版本，等待权限生效
3. 在「ZGame大文本」文件夹 → 添加协作者 → 把该应用加为「可阅读」
   仅开通 scope 不够：应用还必须能看到这个文件夹。

依赖
----
  pip install openpyxl requests
"""

import argparse
import base64
import hashlib
import json
import os
import re
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Optional

import requests
from openpyxl import load_workbook

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ----------------------------- 默认配置 -----------------------------

FEISHU_BASE = "https://open.feishu.cn/open-apis"
DEFAULT_FOLDER_TOKEN = "fldcnvc54hgGywuCsTssbUTlwYg"
DEFAULT_FILE_PATTERN = r"^Zgame\s*ALT.*\.xlsx$"

DEFAULT_SHEET = "Total"
EXPECTED_A1 = "字符串"
BUFF_KEY_RE = re.compile(r"^BUFFAttribute_m(Name|Tips)_(\d+)$")

DEFAULT_GITHUB_OWNER = "PanPerfume"
DEFAULT_GITHUB_REPO = "zgame-buff-data"
DEFAULT_GITHUB_BRANCH = "main"

MIN_XLSX_BYTES = 1024 * 1024  # 小于 1MB 基本可以判定下错文件


class SyncError(RuntimeError):
    pass


# ----------------------------- 飞书 -----------------------------

def feishu_tenant_token(app_id: str, app_secret: str) -> str:
    r = requests.post(
        f"{FEISHU_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=30,
    )
    if r.status_code != 200:
        raise SyncError(f"获取 tenant_access_token 失败：HTTP {r.status_code} {r.text[:300]}")
    data = r.json()
    if data.get("code") != 0 or not data.get("tenant_access_token"):
        raise SyncError(f"获取 tenant_access_token 失败：{data.get('code')} {data.get('msg')}")
    return data["tenant_access_token"]


def _feishu_error(r: requests.Response, action: str) -> SyncError:
    try:
        d = r.json()
        code, msg = d.get("code"), d.get("msg", "")
    except Exception:
        code, msg = r.status_code, (r.text or "")[:300]

    hint = ""
    if code == 99991672:
        hint = (
            "\n→ 应用缺少 drive 权限。打开开放平台申请 drive:drive:readonly 后重新发布版本：\n"
            "  https://open.feishu.cn/app/{app_id}/auth?q=drive:drive:readonly"
        ).format(app_id=os.environ.get("FEISHU_APP_ID", "<APP_ID>"))
    elif code in (1061004, 1061045, 1770001):
        hint = "\n→ 应用还看不到这个文件夹：在文件夹「添加协作者」里把应用加为可阅读。"
    return SyncError(f"{action}失败：{code} {msg}{hint}")


def feishu_list_files(token: str, folder_token: str) -> list:
    files, page_token = [], None
    while True:
        params = {
            "folder_token": folder_token,
            "page_size": 200,
            "order_by": "EditedTime",
            "direction": "DESC",
        }
        if page_token:
            params["page_token"] = page_token
        r = requests.get(
            f"{FEISHU_BASE}/drive/v1/files",
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=30,
        )
        if r.status_code != 200:
            raise _feishu_error(r, "列出文件夹")
        d = r.json()
        if d.get("code") != 0:
            raise _feishu_error(r, "列出文件夹")
        data = d.get("data") or {}
        files.extend(data.get("files") or [])
        if not data.get("has_more"):
            break
        page_token = data.get("next_page_token")
        if not page_token:
            break
    return files


def pick_latest_big_text(files: list, pattern: str) -> dict:
    rx = re.compile(pattern, re.IGNORECASE)
    candidates = [
        f for f in files
        if f.get("type") == "file"
        and rx.search(f.get("name") or "")
        and not (f.get("name") or "").startswith("~$")
    ]
    if not candidates:
        names = [f.get("name") for f in files if f.get("type") == "file"]
        raise SyncError(
            f"文件夹里没有匹配 {pattern} 的文件。当前文件：{names}\n"
            "→ 用 --pattern 调整文件名规则，或检查是否还没上传。"
        )
    candidates.sort(key=lambda f: int(f.get("modified_time") or 0), reverse=True)
    return candidates[0]


def feishu_download(token: str, file_token: str, dest: Path) -> int:
    url = f"{FEISHU_BASE}/drive/v1/files/{file_token}/download"
    with requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        stream=True,
        timeout=300,
    ) as r:
        if r.status_code != 200:
            raise _feishu_error(r, "下载文件")
        ctype = (r.headers.get("Content-Type") or "").lower()
        if "application/json" in ctype:
            raise _feishu_error(r, "下载文件")
        dest.parent.mkdir(parents=True, exist_ok=True)
        size, last_report = 0, time.time()
        with open(dest, "wb") as fh:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if not chunk:
                    continue
                fh.write(chunk)
                size += len(chunk)
                if time.time() - last_report > 3:
                    print(f"  已下载 {size / 1024 / 1024:.1f} MB ...", flush=True)
                    last_report = time.time()
    return size


# ----------------------------- 解析 / 生成 -----------------------------

def safe_text(value) -> str:
    return "" if value is None else str(value)


def sort_buff_rows(rows: list) -> list:
    def sort_key(item):
        key = item.get(EXPECTED_A1, "")
        m = BUFF_KEY_RE.match(key)
        if not m:
            return 999999999, 99, key
        return int(m.group(2)), (0 if m.group(1) == "Name" else 1), key

    return sorted(rows, key=sort_key)


def build_buff_json(xlsx_path: Path, sheet_name: str, source_meta: dict) -> dict:
    """读取 Excel → buff-index.json → 返回统计信息。"""
    print(f"解析：{xlsx_path}（页签 {sheet_name}）")
    wb = load_workbook(str(xlsx_path), read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise SyncError(f"找不到页签 {sheet_name}，当前页签：{wb.sheetnames}")
        ws = wb[sheet_name]
        it = ws.iter_rows(values_only=True)
        try:
            headers_raw = next(it)
        except StopIteration:
            raise SyncError(f"{sheet_name} 页签为空")
        headers = [safe_text(h).strip() for h in headers_raw]
        if not headers or headers[0] != EXPECTED_A1:
            raise SyncError(f"A1 表头应为「{EXPECTED_A1}」，实际是「{headers[0] if headers else ''}」")

        sheet_rows = []
        scanned = 0
        name_rows = tips_rows = 0
        ids = set()

        for row in it:
            scanned += 1
            if not row:
                continue
            key = safe_text(row[0]).strip()
            m = BUFF_KEY_RE.match(key)
            if not m:
                continue
            ids.add(m.group(2))
            if m.group(1) == "Name":
                name_rows += 1
            else:
                tips_rows += 1

            item = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                item[header] = safe_text(row[i]) if i < len(row) else ""
            sheet_rows.append(item)
    finally:
        wb.close()

    if not sheet_rows:
        raise SyncError("没有解析到任何 BUFFAttribute_mName/mTips 行，检查表结构是否变了。")

    sheet_rows = sort_buff_rows(sheet_rows)
    index_json = json.dumps({"Sheet1": sheet_rows}, ensure_ascii=False, indent=4)
    content_hash = hashlib.sha256(index_json.encode("utf-8")).hexdigest()[:12]

    now = datetime.now()
    meta = {
        "version": f"{now.strftime('%Y%m%d-%H%M%S')}-{content_hash}",
        "updatedAt": now.strftime("%Y-%m-%d %H:%M:%S"),
        "dataUrl": "buff-index.json",
        "sourceFile": source_meta.get("name") or xlsx_path.name,
        "sheetName": sheet_name,
        "rowCount": len(sheet_rows),
        "buffIdCount": len(ids),
        "nameRowCount": name_rows,
        "tipsRowCount": tips_rows,
        "format": "pretty",
    }
    if source_meta.get("token"):
        meta["sourceToken"] = source_meta["token"]
    if source_meta.get("modified"):
        meta["sourceModified"] = source_meta["modified"]

    return {
        "index_json": index_json,
        "meta_json": json.dumps(meta, ensure_ascii=False, indent=4),
        "stats": {
            "scanned": scanned,
            "rows": len(sheet_rows),
            "ids": len(ids),
            "name": name_rows,
            "tips": tips_rows,
            "version": meta["version"],
            "index_bytes": len(index_json.encode("utf-8")),
        },
    }


def atomic_write_text(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile("w", delete=False, encoding="utf-8",
                            dir=str(path.parent), suffix=".tmp") as f:
        f.write(text)
        tmp = Path(f.name)
    tmp.replace(path)


# ----------------------------- GitHub -----------------------------

def github_headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_upload(token: str, owner: str, repo: str, branch: str,
                  local_path: Path, repo_path: str):
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{repo_path}"
    local_bytes = local_path.read_bytes()

    sha = None
    r = requests.get(url, headers=github_headers(token),
                     params={"ref": branch}, timeout=30)
    if r.status_code == 200:
        d = r.json()
        sha = d.get("sha")
        remote_content = d.get("content")
        if remote_content and d.get("encoding") == "base64":
            if base64.b64decode(remote_content) == local_bytes:
                print(f"  内容未变化，跳过：{repo_path}")
                return False
    elif r.status_code != 404:
        raise SyncError(f"查询 GitHub 文件失败：HTTP {r.status_code} {r.text[:300]}")

    body = {
        "message": f"Update {repo_path} ({datetime.now():%Y-%m-%d %H:%M})",
        "content": base64.b64encode(local_bytes).decode("ascii"),
        "branch": branch,
    }
    if sha:
        body["sha"] = sha

    for attempt in range(2):
        resp = requests.put(url, headers=github_headers(token), json=body, timeout=120)
        if resp.status_code in (200, 201):
            commit = (resp.json().get("commit") or {}).get("sha", "")[:10]
            print(f"  上传完成：{repo_path}  commit={commit}")
            return True
        if resp.status_code == 409 and attempt == 0:
            # 并发更新：重新取 sha 再试一次
            time.sleep(2)
            rr = requests.get(url, headers=github_headers(token),
                              params={"ref": branch}, timeout=30)
            if rr.status_code == 200:
                body["sha"] = rr.json().get("sha")
                continue
        raise SyncError(f"上传 GitHub 失败：{repo_path} HTTP {resp.status_code} {resp.text[:300]}")
    return False


# ----------------------------- 本地兜底 -----------------------------

def find_latest_local(pattern: str, extra_dirs=None) -> Path:
    dirs = [Path.home() / "Desktop", Path.home() / "Downloads",
            Path("C:/Users/admin/Desktop")]
    if extra_dirs:
        dirs.extend(Path(d) for d in extra_dirs)
    rx = re.compile(pattern, re.IGNORECASE)
    hits = []
    for d in dirs:
        if not d.exists():
            continue
        for p in d.glob("*.xlsx"):
            if rx.search(p.name) and not p.name.startswith("~$"):
                hits.append(p)
    if not hits:
        raise SyncError(f"在 {[str(d) for d in dirs]} 里没找到匹配 {pattern} 的 xlsx。")
    hits.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return hits[0]


# ----------------------------- 主流程 -----------------------------

def load_env_file():
    """
    读取脚本同目录的 config.env（KEY=VALUE），已存在的环境变量优先。
    这样双击 run_sync.bat 即可运行，无需手动 set 环境变量。
    """
    p = Path(__file__).resolve().with_name("config.env")
    if not p.exists():
        return
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        k, v = k.strip(), v.strip().strip('"').strip("'")
        if k and k not in os.environ:
            os.environ[k] = v


def main():
    load_env_file()
    ap = argparse.ArgumentParser(description="飞书大文本 → Buff JSON → GitHub")
    ap.add_argument("--local", type=str, help="直接用本地 xlsx，不访问飞书")
    ap.add_argument("--auto-local", action="store_true",
                    help="在 Desktop/Downloads 自动找最新 Zgame ALT*.xlsx")
    ap.add_argument("--pattern", default=DEFAULT_FILE_PATTERN,
                    help=f"文件名正则（默认 {DEFAULT_FILE_PATTERN}）")
    ap.add_argument("--folder-token", default=os.environ.get("FEISHU_FOLDER_TOKEN", DEFAULT_FOLDER_TOKEN))
    ap.add_argument("--sheet", default=DEFAULT_SHEET)
    ap.add_argument("--output-dir", default="./buff-data-out",
                    help="生成的 JSON 存放目录（默认 ./buff-data-out）")
    ap.add_argument("--dry-run", action="store_true", help="只生成，不上传")
    ap.add_argument("--skip-upload", action="store_true", help="同 --dry-run")
    ap.add_argument("--force", action="store_true", help="内容未变化也重新生成并上传")
    ap.add_argument("--keep", action="store_true", help="保留下载的 xlsx")
    args = ap.parse_args()

    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path: Optional[Path] = None
    temp_xlsx: Optional[Path] = None
    source_meta = {}

    # ---- 1. 取源文件 ----
    if args.local:
        xlsx_path = Path(args.local).expanduser()
        if not xlsx_path.exists():
            raise SyncError(f"本地文件不存在：{xlsx_path}")
        source_meta = {"name": xlsx_path.name}
        print(f"使用本地文件：{xlsx_path}")

    elif args.auto_local:
        xlsx_path = find_latest_local(args.pattern)
        source_meta = {"name": xlsx_path.name}
        print(f"自动选中本地最新：{xlsx_path}")

    else:
        app_id = os.environ.get("FEISHU_APP_ID")
        app_secret = os.environ.get("FEISHU_APP_SECRET")
        if not app_id or not app_secret:
            raise SyncError(
                "未设置 FEISHU_APP_ID / FEISHU_APP_SECRET。\n"
                "可用 --local <xlsx> 或 --auto-local 先跑通；\n"
                "或设置环境变量后重试。"
            )

        print("获取飞书 token ...")
        token = feishu_tenant_token(app_id, app_secret)
        print(f"列出文件夹 {args.folder_token} ...")
        files = feishu_list_files(token, args.folder_token)
        picked = pick_latest_big_text(files, args.pattern)
        modified = datetime.fromtimestamp(int(picked.get("modified_time") or 0))
        print(f"选中：{picked['name']}（修改于 {modified:%Y-%m-%d %H:%M}）")

        temp_xlsx = Path(tempfile.gettempdir()) / f"viz_{picked['token']}.xlsx"
        print("下载中 ...")
        size = feishu_download(token, picked["token"], temp_xlsx)
        if size < MIN_XLSX_BYTES:
            raise SyncError(f"下载文件只有 {size} 字节，疑似不是完整 Excel。")
        print(f"下载完成：{size / 1024 / 1024:.1f} MB")

        xlsx_path = temp_xlsx
        source_meta = {
            "name": picked["name"],
            "token": picked["token"],
            "modified": f"{modified:%Y-%m-%d %H:%M:%S}",
        }

    if xlsx_path.stat().st_size < MIN_XLSX_BYTES:
        raise SyncError(f"源文件只有 {xlsx_path.stat().st_size} 字节，疑似损坏。")

    # ---- 2. 生成 ----
    built = build_buff_json(xlsx_path, args.sheet, source_meta)
    s = built["stats"]

    index_file = out_dir / "buff-index.json"
    meta_file = out_dir / "buff-meta.json"

    # 内容未变化时保留原 meta（version 不变），避免空提交和让插件重复下载
    unchanged = False
    if not args.force and index_file.exists() and meta_file.exists():
        try:
            unchanged = index_file.read_text(encoding="utf-8") == built["index_json"]
        except OSError:
            unchanged = False

    if unchanged:
        print("数据内容未变化，保留现有 buff-meta.json（版本号不变）。")
    else:
        atomic_write_text(index_file, built["index_json"])
        atomic_write_text(meta_file, built["meta_json"])

    print("")
    print("=" * 62)
    print(f"扫描行数      ：{s['scanned']:,}")
    print(f"导出 Buff 行  ：{s['rows']:,}")
    print(f"Buff ID 数量  ：{s['ids']:,}")
    print(f"Name / Tips   ：{s['name']:,} / {s['tips']:,}")
    print(f"JSON 大小     ：{s['index_bytes'] / 1024 / 1024:.2f} MB")
    print(f"版本          ：{s['version']}")
    print(f"输出目录      ：{out_dir}")
    print("=" * 62)

    if unchanged:
        print("已跳过上传（内容未变化）。")
    elif args.dry_run or args.skip_upload:
        print("已跳过上传（--dry-run/--skip-upload）。")
    else:
        gh_token = os.environ.get("GITHUB_TOKEN")
        if not gh_token:
            raise SyncError("未设置 GITHUB_TOKEN，无法上传；如需只生成请加 --dry-run。")
        owner = os.environ.get("GITHUB_OWNER", DEFAULT_GITHUB_OWNER)
        repo = os.environ.get("GITHUB_REPO", DEFAULT_GITHUB_REPO)
        branch = os.environ.get("GITHUB_BRANCH", DEFAULT_GITHUB_BRANCH)
        print(f"上传到 {owner}/{repo}@{branch} ...")
        github_upload(gh_token, owner, repo, branch, index_file, "buff-index.json")
        github_upload(gh_token, owner, repo, branch, meta_file, "buff-meta.json")
        print(f"Raw：https://raw.githubusercontent.com/{owner}/{repo}/{branch}/buff-meta.json")

    if temp_xlsx and not args.keep:
        try:
            temp_xlsx.unlink()
        except OSError:
            pass


if __name__ == "__main__":
    try:
        main()
    except SyncError as e:
        print(f"\n[失败] {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(130)
