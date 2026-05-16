#!/usr/bin/env python3
"""
OneDrive → GitHub 知识库同步脚本
使用 MSAL 客户端凭证模式，无需手动刷新 Token
使用文件夹 ID 替代路径，更稳定可靠
"""

import os
import sys
import json
import hashlib
from pathlib import Path

import requests
from msal import ConfidentialClientApplication

# ── 配置 ──────────────────────────────────────────────────
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

# 要同步的 OneDrive 文件夹 ID（PC_Knowledges_Library）
ROOT_FOLDER_ID = "016EIT54A47KP7N2LYZFALW4IIGA2Y7HLJ"
# 使用 GITHUB_WORKSPACE（GitHub Actions 自动设置）或默认路径
GITHUB_WORKSPACE = os.environ.get("GITHUB_WORKSPACE", "/github/workspace")
LOCAL_REPO = Path(os.environ.get("LOCAL_REPO_PATH", GITHUB_WORKSPACE))
SYNC_LOG = LOCAL_REPO / ".github" / "sync_status.json"

SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# OneDrive Drive ID
DRIVE_ID = os.environ.get("AZURE_DRIVE_ID", "b!K4wK9ZoVNkmnE1bHe97JzPppa1Pc24lKu4INPlht8NhY1WKODb6FRIGPv1cAeN3_")


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
    print(f"✅ Token 获取成功 (expires: {result.get('expires_in', '?')}s)")
    return result["access_token"]


# ── 读取 OneDrive 文件结构 ────────────────────────────────
def list_onedrive_files(token, folder_id=ROOT_FOLDER_ID):
    """递归列出 OneDrive 文件夹下所有文件（基于文件夹 ID，不用路径）"""
    headers = {"Authorization": f"Bearer {token}"}
    files = []

    def recurse(item_id, relative_path=""):
        url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{item_id}/children"
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print(f"  ⚠️ 无法读取 {relative_path or '(根目录)'}: {resp.status_code} {resp.text[:200]}")
            return

        for item in resp.json().get("value", []):
            name = item["name"]
            item_path = f"{relative_path}/{name}" if relative_path else name

            if item.get("folder"):
                # 排除系统目录
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
                    "downloadUrl": item.get("@microsoft.graph.downloadUrl", ""),
                })

        # 处理分页
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
                        "downloadUrl": item.get("@microsoft.graph.downloadUrl", ""),
                    })
            next_link = resp2.json().get("@odata.nextLink")

    recurse(folder_id)
    print(f"  📄 共发现 {len(files)} 个文件")
    return files


# ── 计算文件哈希 ──────────────────────────────────────────
def file_hash(path):
    if not path.exists():
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ── 下载文件到仓库 ────────────────────────────────────────
def download_file(token, download_url, local_path):
    local_path.parent.mkdir(parents=True, exist_ok=True)
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(download_url, headers=headers, stream=True)
    if resp.status_code != 200:
        print(f"  ⚠️ 下载失败 {local_path.name}: HTTP {resp.status_code}")
        return False
    with open(local_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
    return True


# ── 主流程 ────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("📚 知识库同步：OneDrive → GitHub")
    print(f"📂 文件夹 ID: {ROOT_FOLDER_ID}")
    print(f"📂 本地仓库路径: {LOCAL_REPO}")
    print("=" * 50)

    # 1. 获取 Token
    token = get_token()

    # 2. 读取 OneDrive 文件列表
    print("\n🔍 扫描 OneDrive 文件...")
    remote_files = list_onedrive_files(token)

    if not remote_files:
        print("  ⚠️ OneDrive 中没有文件，请检查文件夹 ID")
        sys.exit(0)

    # 3. 逐文件同步
    print("\n📥 同步文件...")
    downloaded = 0
    skipped = 0
    failed = 0

    for rf in remote_files:
        local = LOCAL_REPO / rf["path"]

        # 跳过临时文件
        if local.name.startswith("~$") or local.suffix in (".tmp", ".temp", ".lnk"):
            continue

        # 按大小判断是否需要更新
        try:
            if local.exists() and local.stat().st_size == rf["size"]:
                skipped += 1
                continue
        except OSError:
            pass

        # 下载
        if rf.get("downloadUrl"):
            ok = download_file(token, rf["downloadUrl"], local)
            if ok:
                downloaded += 1
                print(f"  ✅ {rf['path']}")
            else:
                failed += 1
                print(f"  ❌ {rf['path']}")
        else:
            # 没有 downloadUrl 时，用 content 端点
            headers = {"Authorization": f"Bearer {token}"}
            url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{rf['id']}/content"
            resp = requests.get(url, headers=headers)
            if resp.status_code == 200:
                local.parent.mkdir(parents=True, exist_ok=True)
                local.write_bytes(resp.content)
                downloaded += 1
                print(f"  ✅ {rf['path']}")
            else:
                failed += 1
                print(f"  ❌ {rf['path']} (HTTP {resp.status_code})")

    # 4. 清理远程已删除的文件
    print("\n🧹 检查已删除的文件...")
    remote_paths = {rf["path"] for rf in remote_files}

    # 仓库基础设施文件——不在 OneDrive 中，但必须保留
    PROTECTED_FILES = {"README.md", "LICENSE", ".gitignore"}
    PROTECTED_PREFIXES = (".git", ".github")

    deleted = 0
    for local_file in LOCAL_REPO.rglob("*"):
        if local_file.is_file():
            relative = str(local_file.relative_to(LOCAL_REPO))
            if any(relative.startswith(p) for p in PROTECTED_PREFIXES):
                continue
            if relative in PROTECTED_FILES:
                continue
            if relative not in remote_paths:
                local_file.unlink()
                deleted += 1
                print(f"  🗑️ {relative}")

    # 5. 保存同步状态
    status = {
        "last_sync": os.popen("date -u '+%Y-%m-%dT%H:%M:%SZ'").read().strip(),
        "downloaded": downloaded,
        "skipped": skipped,
        "failed": failed,
        "deleted": deleted,
        "total_remote": len(remote_files),
    }
    SYNC_LOG.parent.mkdir(parents=True, exist_ok=True)
    with open(SYNC_LOG, "w") as f:
        json.dump(status, f, indent=2)

    # 6. 摘要
    print("\n" + "=" * 50)
    print(f"📊 同步完成")
    print(f"  ✅ 下载: {downloaded}")
    print(f"  ⏭️  跳过(无变化): {skipped}")
    print(f"  ❌ 失败: {failed}")
    print(f"  🗑️  删除: {deleted}")
    print(f"  📄 远程文件数: {len(remote_files)}")
    print("=" * 50)

    if failed > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
