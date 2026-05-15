# 📚 Knowledge Library

自动同步 OneDrive 的知识库到 GitHub，由《暝烁》管理。

## 工作流程

```
主子在电脑上编辑笔记/资料
        ↓ OneDrive 自动同步（主子电脑已有）
OneDrive 云端
        ↓ GitHub Action（每周六 08:00 自动运行）
获取 Graph API Token → 读取 OneDrive 文件 → 同步到 GitHub
        ↓
等待下一次同步...
```

## 配置

OneDrive 路径通过 GitHub Secret `ONEDRIVE_ROOT` 配置，默认为 `/KnowledgeLibrary`。
如需修改，在仓库 Settings → Secrets → Actions 中添加 `ONEDRIVE_ROOT`。

## 手动触发

在 GitHub 仓库的 Actions 页面，点 `Sync from OneDrive` → `Run workflow` 即可。
