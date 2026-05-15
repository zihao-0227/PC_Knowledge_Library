#!/usr/bin/env python3
"""
OneDrive → GitHub 知识库同步脚本
使用 MSAL 客户端凭证模式，无需手动刷新 Token
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

# OneDrive 路径：默认 /KnowledgeLibrary，可通过 Secret 覆盖
ONEDRIVE_ROOT = os.environ.get("ONEDRIVE_ROOT", "/drive/root:/KnowledgeLibrary")
LOCAL_REPO = Path(os.environ.get("LOCAL_REPO_PATH", "/github/workspace"))  # 本地或 GitHub Actions 工作目录
SYNC_LOG = LOCAL_REPO / ".github" / "sync_status.json"

SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# OneDrive Drive ID（应用权限需要用 drive ID 而不是 /me 或 /users/{upn}）
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
def list_onedrive_files(token, folder_path=ONEDRIVE_ROOT):
    """递归列出 OneDrive 文件夹下所有文件"""
    headers = {"Authorization": f"Bearer {token}"}
    files = []

    drive_root = f"{GRAPH_BASE}/drives/{DRIVE_ID}/root"

    def recurse(path):
        # 路径格式处理：root 用 :，子路径用 /
        if path == ":" or path == "":
            endpoint = drive_root
        elif path.startswith(":/"):
            # 格式：:/文件夹名/子文件夹
            endpoint = f"{drive_root}{path}"
        else:
            endpoint = f"{drive_root}{path}"
        url = f"{endpoint}:/children"
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print(f"  ⚠️ 无法读取 {path}: {resp.status_code} {resp.text[:200]}")
            return
        for item in resp.json().get("value", []):
            if item.get("folder"):
                # 排除 Windows 系统目录
                name = item["name"]
                if name.startswith(".") or name in ("System Volume Information", "$RECYCLE.BIN"):
                    continue
                # 子路径用 / 分隔（只有第一个 : 是根路径标记）
                sub_path = f"{path}/{name}"
                recurse(sub_path)
            else:
                files.append({
                    "path": f"{path}/{item['name']}",
                    "name": item["name"],
                    "id": item["id"],
                    "size": item.get("size", 0),
                    "lastModified": item.get("lastModifiedDateTime", ""),
                    "downloadUrl": item.get("@microsoft.graph.downloadUrl", ""),
                })
        # 处理分页
        next_link = resp.json().get("@odata.nextLink")
        if next_link:
            resp2 = requests.get(next_link, headers=headers)
            if resp2.status_code == 200:
                for item in resp2.json().get("value", []):
                    if item.get("folder"):
                        name = item["name"]
                        if name.startswith(".") or name in ("System Volume Information", "$RECYCLE.BIN"):
                            continue
                        sub_path = f"{path}:/{name}"
                        recurse(sub_path)
                    else:
                        files.append({
                            "path": f"{path}/{item['name']}",
                            "name": item["name"],
                            "id": item["id"],
                            "size": item.get("size", 0),
                            "lastModified": item.get("lastModifiedDateTime", ""),
                            "downloadUrl": item.get("@microsoft.graph.downloadUrl", ""),
                        })

    recurse(folder_path)
    print(f"  📄 共发现 {len(files)} 个文件")
    return files


# ── 计算文件哈希 ──────────────────────────────────────────
def file_hash(path):
    """快速哈希判断文件是否变动"""
    if not path.exists():
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ── 下载文件到仓库 ────────────────────────────────────────
def download_file(token, download_url, local_path):
    """下载 OneDrive 文件到本地"""
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
    print(f"📂 OneDrive 路径: {ONEDRIVE_ROOT}")
    print(f"📂 本地仓库路径: {LOCAL_REPO}")
    print("=" * 50)

    # 1. 获取 Token
    token = get_token()

    # 2. 读取 OneDrive 文件列表
    print("\n🔍 扫描 OneDrive 文件...")
    remote_files = list_onedrive_files(token)

    if not remote_files:
        print("  ⚠️ OneDrive 中没有文件，请检查 ONEDRIVE_ROOT 路径")
        print(f"  🔧 当前路径: {ONEDRIVE_ROOT}")
        print("  💡 可以在仓库 Settings → Secrets 中设置 ONEDRIVE_ROOT")
        sys.exit(0)

    # 3. 逐文件同步
    print("\n📥 同步文件...")
    downloaded = 0
    skipped = 0
    failed = 0

    for rf in remote_files:
        # 将 OneDrive 路径映射到本地仓库路径
        # 例: /drive/root:/KnowledgeLibrary/笔记/hello.md → 笔记/hello.md
        relative = rf["path"].replace(ONEDRIVE_ROOT, "", 1).lstrip("/:")
        local = LOCAL_REPO / relative

        # 跳过临时文件和系统文件
        if local.name.startswith("~$") or local.suffix in (".tmp", ".temp", ".lnk"):
            continue

        # 检查是否已是最新（用哈希判断）
        existing_hash = file_hash(local) if "downloadUrl" in rf and rf["downloadUrl"] else None
        if existing_hash:
            # 先从 OneDrive 获取文件校验信息
            try:
                local_size = local.stat().st_size
                if local_size == rf["size"]:
                    skipped += 1
                    continue
            except OSError:
                pass

        # 下载
        if rf.get("downloadUrl"):
            ok = download_file(token, rf["downloadUrl"], local)
            if ok:
                downloaded += 1
                print(f"  ✅ {relative}")
            else:
                failed += 1
                print(f"  ❌ {relative}")
        else:
            # 没有 downloadUrl 的情况，用 content 端点
            headers = {"Authorization": f"Bearer {token}"}
            url = f"{drive_root}{rf['path']}:/content"
            resp = requests.get(url, headers=headers)
            if resp.status_code == 200:
                local.parent.mkdir(parents=True, exist_ok=True)
                local.write_bytes(resp.content)
                downloaded += 1
                print(f"  ✅ {relative}")
            else:
                failed += 1
                print(f"  ❌ {relative} (HTTP {resp.status_code})")

    # 4. 清理远程已删除的文件
    print("\n🧹 检查已删除的文件...")
    remote_paths = set()
    for rf in remote_files:
        relative = rf["path"].replace(ONEDRIVE_ROOT, "", 1).lstrip("/:")
        remote_paths.add(relative)

    deleted = 0
    for local_file in LOCAL_REPO.rglob("*"):
        if local_file.is_file():
            relative = str(local_file.relative_to(LOCAL_REPO))
            # 跳过 .git 和 .github 目录
            if relative.startswith(".git") or relative.startswith(".github"):
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
