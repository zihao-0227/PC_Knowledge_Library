#!/usr/bin/env python3
"""
GitHub → OneDrive 知识库反向同步脚本
当 GitHub 仓库有新的推送时，自动将变更同步到 OneDrive
"""
import os
import sys
import json
import hashlib
import time
from pathlib import Path

import requests
from msal import ConfidentialClientApplication

# ── 配置 ──────────────────────────────────────────────────
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

ROOT_FOLDER_ID = "012NGUUTGUIEX7XCQZWJG2YYXTHNKB4T35"
DRIVE_ID = os.environ["AZURE_DRIVE_ID"]

# GitHub Actions 自动设置的工作目录
LOCAL_REPO = Path(os.environ.get("GITHUB_WORKSPACE", "/github/workspace"))
SYNC_LOG = LOCAL_REPO / ".github" / "sync_status.json"

SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# 上传文件大小限制：4MB 以下用小文件上传，以上用分片上传
CHUNK_SIZE_LIMIT = 4 * 1024 * 1024


# ── 认证 ──────────────────────────────────────────────────
def get_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_silent(SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in result:
        print(f"❌ 获取 Token 失败: {result.get('error_description', result)}")
        sys.exit(1)
    print(f"✅ Token 获取成功")
    return result["access_token"]


# ── OneDrive 文件操作 ─────────────────────────────────────
def list_onedrive_files(token, folder_id=ROOT_FOLDER_ID):
    """递归列出 OneDrive 文件夹下所有文件（和正向同步保持一致）"""
    headers = {"Authorization": f"Bearer {token}"}
    files = []

    def recurse(item_id, relative_path=""):
        url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{item_id}/children"
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print(f"  ⚠️ 无法读取 {relative_path or '(根目录)'}: HTTP {resp.status_code}")
            return

        for item in resp.json().get("value", []):
            name = item["name"]
            item_path = f"{relative_path}/{name}" if relative_path else name

            if item.get("folder"):
                if name.startswith(".") or name in ("System Volume Information", "$RECYCLE.BIN"):
                    continue
                recurse(item["id"], item_path)
            else:
                files.append({
                    "path": item_path,
                    "name": name,
                    "id": item["id"],
                    "size": item.get("size", 0),
                    "lastModified": item.get("lastModifiedDateTime", ""),
                    "eTag": item.get("@odata.etag", ""),
                })

        # 分页处理
        next_link = resp.json().get("@odata.nextLink")
        while next_link:
            resp2 = requests.get(next_link, headers=headers)
            if resp2.status_code != 200:
                break
            for item in resp2.json().get("value", []):
                name = item["name"]
                item_path = f"{relative_path}/{name}" if relative_path else name
                if item.get("folder"):
                    if name.startswith(".") or name in ("System Volume Information", "$RECYCLE.BIN"):
                        continue
                    recurse(item["id"], item_path)
                else:
                    files.append({
                        "path": item_path,
                        "name": name,
                        "id": item["id"],
                        "size": item.get("size", 0),
                        "lastModified": item.get("lastModifiedDateTime", ""),
                        "eTag": item.get("@odata.etag", ""),
                    })
            next_link = resp2.json().get("@odata.nextLink")

    recurse(folder_id)
    print(f"  📄 OneDrive 中现有 {len(files)} 个文件")
    return files


def find_or_create_folder(token, parent_id, folder_name):
    """在 OneDrive 中查找或创建文件夹"""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # 先查找
    url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{parent_id}/children"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        for item in resp.json().get("value", []):
            if item.get("folder") and item["name"] == folder_name:
                return item["id"]
    
    # 没找到，创建
    url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{parent_id}/children"
    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail"
    }
    resp = requests.post(url, headers=headers, json=body)
    if resp.status_code in (200, 201):
        return resp.json()["id"]
    
    # 如果冲突了（已存在但刚才没查到），再查一次
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        for item in resp.json().get("value", []):
            if item.get("folder") and item["name"] == folder_name:
                return item["id"]
    
    print(f"  ❌ 无法创建文件夹 {folder_name}: HTTP {resp.status_code}")
    return None


def ensure_folder_path(token, path_parts):
    """确保 OneDrive 中的文件夹路径存在，返回最终文件夹 ID"""
    current_id = ROOT_FOLDER_ID
    for part in path_parts:
        if not part:
            continue
        current_id = find_or_create_folder(token, current_id, part)
        if not current_id:
            return None
    return current_id


def upload_small_file(token, local_path, parent_id, filename):
    """上传小文件（<4MB）"""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{parent_id}:/{filename}:/content"
    
    with open(local_path, "rb") as f:
        content = f.read()
    
    resp = requests.put(url, headers=headers, data=content)
    if resp.status_code in (200, 201):
        return True
    else:
        print(f"    ❌ 上传失败: HTTP {resp.status_code}")
        return False


def upload_large_file(token, local_path, parent_id, filename):
    """上传大文件（>=4MB），使用分片上传"""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # 创建上传会话
    url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{parent_id}:/{filename}:/createUploadSession"
    body = {
        "@microsoft.graph.conflictBehavior": "replace",
        "name": filename
    }
    resp = requests.post(url, headers=headers, json=body)
    if resp.status_code != 200:
        print(f"    ❌ 创建上传会话失败: HTTP {resp.status_code}")
        return False
    
    upload_url = resp.json()["uploadUrl"]
    file_size = os.path.getsize(local_path)
    
    # 分片上传
    with open(local_path, "rb") as f:
        chunk_size = 5 * 1024 * 1024  # 5MB 分片
        uploaded = 0
        while uploaded < file_size:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            end = min(uploaded + len(chunk) - 1, file_size - 1)
            
            chunk_headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {uploaded}-{end}/{file_size}"
            }
            resp = requests.put(upload_url, headers=chunk_headers, data=chunk)
            uploaded += len(chunk)
    
    return True


def upload_file_to_onedrive(token, local_path, relative_path):
    """上传文件到 OneDrive"""
    path_parts = relative_path.split("/")
    filename = path_parts[-1]
    folder_parts = path_parts[:-1]
    
    print(f"  📤 {relative_path}...", end=" ")
    
    # 确保文件夹存在
    parent_id = ensure_folder_path(token, folder_parts)
    if not parent_id:
        print("❌ 无法创建文件夹路径")
        return False
    
    # 按文件大小选择上传方式
    file_size = os.path.getsize(local_path)
    if file_size < CHUNK_SIZE_LIMIT:
        ok = upload_small_file(token, local_path, parent_id, filename)
    else:
        ok = upload_large_file(token, local_path, parent_id, filename)
    
    if ok:
        print("✅")
    return ok


def delete_onedrive_file(token, item_id, path_name):
    """删除 OneDrive 中的文件"""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{item_id}"
    resp = requests.delete(url, headers=headers)
    if resp.status_code in (200, 204):
        print(f"  🗑️ {path_name}")
        return True
    else:
        print(f"  ⚠️ 删除失败 {path_name}: HTTP {resp.status_code}")
        return False


# ── 本地文件扫描 ──────────────────────────────────────────
def list_local_files():
    """扫描 GitHub 仓库中的所有文件（排除 .git、.github 和 Quartz 引擎文件）"""
    PROTECTED_PREFIXES = (".git", ".github", "quartz/", "node_modules/", "public/")
    PROTECTED_FILES = {
        "README.md", "LICENSE", ".gitignore",
        "package.json", "package-lock.json", "quartz.config.yaml",
        "tsconfig.json", "globals.d.ts", "index.d.ts",
        "quartz.ts", ".node-version", ".npmrc",
    }
    
    files = []
    for local_file in sorted(LOCAL_REPO.rglob("*")):
        if not local_file.is_file():
            continue
        relative = str(local_file.relative_to(LOCAL_REPO))
        if any(relative.startswith(p) for p in PROTECTED_PREFIXES):
            continue
        if relative in PROTECTED_FILES:
            continue
        files.append({
            "path": relative,
            "size": local_file.stat().st_size,
        })
    
    print(f"  📄 GitHub 仓库中有 {len(files)} 个文件")
    return files


# ── 文件哈希 ──────────────────────────────────────────────
def file_sha256(path):
    if not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ── 主流程 ────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("📚 知识库反向同步：GitHub → OneDrive")
    print(f"📂 OneDrive 文件夹 ID: {ROOT_FOLDER_ID}")
    print(f"📂 本地仓库: {LOCAL_REPO}")
    print("=" * 50)

    # 1. 获取 Token
    token = get_token()

    # 2. 读取 OneDrive 文件列表（建立 path → id 映射）
    print("\n🔍 扫描 OneDrive...")
    od_files = list_onedrive_files(token)
    od_by_path = {f["path"]: f for f in od_files}
    od_paths = set(od_by_path.keys())

    # 3. 读取 GitHub 仓库文件列表
    print("\n🔍 扫描 GitHub 仓库...")
    local_files = list_local_files()
    local_paths = {f["path"] for f in local_files}
    local_by_path = {f["path"]: f for f in local_files}

    # 4. 计算差异
    # 需上传的：GitHub 有但 OneDrive 没有，或大小不同
    uploads = []
    for lf in local_files:
        if lf["path"] not in od_paths:
            uploads.append(lf["path"])
        elif od_by_path[lf["path"]]["size"] != lf["size"]:
            uploads.append(lf["path"])
    
    # 需删除的：OneDrive 有但 GitHub 没有
    deletions = []
    for od_path in od_paths:
        if od_path not in local_paths:
            deletions.append(od_path)
    
    print(f"\n📊 差异统计：")
    print(f"  📤 需上传: {len(uploads)} 个文件")
    print(f"  🗑️  需删除: {len(deletions)} 个文件")

    # 5. 执行上传
    if uploads:
        print(f"\n📤 上传到 OneDrive...")
        upload_ok = 0
        upload_fail = 0
        for rel_path in uploads:
            local_path = LOCAL_REPO / rel_path
            if local_path.exists():
                ok = upload_file_to_onedrive(token, local_path, rel_path)
                if ok:
                    upload_ok += 1
                else:
                    upload_fail += 1
            else:
                print(f"  ⚠️ 本地文件不存在: {rel_path}")
                upload_fail += 1
        print(f"  上传完成: ✅ {upload_ok} | ❌ {upload_fail}")
    else:
        print(f"\n📤 无需上传任何文件")

    # 6. 执行删除
    if deletions:
        print(f"\n🗑️  删除 OneDrive 中已移除的文件...")
        del_ok = 0
        del_fail = 0
        for del_path in deletions:
            if del_path in od_by_path:
                ok = delete_onedrive_file(token, od_by_path[del_path]["id"], del_path)
                if ok:
                    del_ok += 1
                else:
                    del_fail += 1
        print(f"  删除完成: ✅ {del_ok} | ❌ {del_fail}")
    else:
        print(f"\n🗑️  无需删除任何文件")

    # 7. 保存状态
    status = {
        "last_sync": os.popen("date -u '+%Y-%m-%dT%H:%M:%SZ'").read().strip(),
        "uploaded": len(uploads),
        "deleted": len(deletions),
    }
    SYNC_LOG.parent.mkdir(parents=True, exist_ok=True)
    with open(SYNC_LOG, "w") as f:
        json.dump(status, f, indent=2)

    # 8. 摘要
    print("\n" + "=" * 50)
    print("📊 反向同步完成")
    print(f"  📤 上传: {len(uploads)}")
    print(f"  🗑️  删除: {len(deletions)}")
    print("=" * 50)


if __name__ == "__main__":
    main()
