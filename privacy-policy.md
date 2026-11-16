# Visualizer 隐私政策 / Privacy Policy

**生效日期：2026-11-16**
**开发者：Visualizer Team**
**联系方式：2581832218@qq.com**
**适用版本：Visualizer 6.1.1**

---

## 中文

### 1. 我们收集什么

Visualizer（以下简称“本扩展”）**不收集、不上传、不出售任何个人数据**。

具体而言：

- 本扩展不会收集你的姓名、邮箱、账号、浏览历史或设备标识；
- 你在网页中选中的文本只在你本机浏览器内解析和显示，不会发送到任何服务器；
- 本扩展不包含统计埋点、广告 SDK 或任何第三方追踪代码。

### 2. 网络请求

本扩展仅向以下地址发起请求，用于获取公开的游戏数据文件（JSON）：

- `https://raw.githubusercontent.com/PanPerfume/zgame-buff-data/...`

这些请求不携带你的个人身份信息，也不包含你浏览的页面内容。请求得到的公开数据缓存在你本机浏览器存储中，用于减少重复下载。

### 3. 本地存储

本扩展使用 `chrome.storage` 保存以下内容：

- 公开 Buff 数据的本地缓存；
- 你的界面偏好（语言、夜间模式、布局等）。

以上内容均保存在你本机，不会上传。

### 4. 权限用途

| 权限 | 用途 |
|---|---|
| `contextMenus` | 在用户选中文本时提供“Visualize”右键菜单入口 |
| `sidePanel` | 在浏览器侧边栏展示 Source / Target 双栏预览界面 |
| `windows` | 支持将预览面板打开为独立窗口 |
| `system.display` | 计算独立窗口在当前显示器上的尺寸与位置 |
| `storage` | 保存界面偏好与公开数据的本地缓存 |
| `unlimitedStorage` | 公开 Buff 数据体积可能超过默认存储配额，缓存需要更多空间 |
| `alarms` | 每日检查一次公开数据是否有更新 |
| `host_permissions: https://raw.githubusercontent.com/*` | 从该公共地址拉取公开的游戏 Buff 数据文件（JSON），不涉及任何用户数据 |
| `content_scripts: <all_urls>` | 核心功能需要读取用户在任意网页上选中的文本并在本地渲染，无法预先限定站点 |

### 5. 儿童隐私

本扩展不面向 13 岁以下儿童，也不收集任何年龄相关信息。

### 6. 政策变更

如本政策发生变更，我们会在本页面更新生效日期，并在扩展的版本说明中注明。

### 7. 联系

如有任何隐私相关问题，请联系：2581832218@qq.com

---

## English

### 1. What we collect

Visualizer (the "Extension") **does not collect, transmit, or sell any personal data**.

- We do not collect names, email addresses, accounts, browsing history, or device identifiers.
- Text you select on a web page is parsed and rendered locally in your browser and is never sent to a server.
- The Extension contains no analytics, advertising SDKs, or third-party tracking code.

### 2. Network requests

The Extension makes requests only to the following public endpoint to fetch open game data (JSON):

- `https://raw.githubusercontent.com/PanPerfume/zgame-buff-data/...`

These requests contain no personal identifiers and no page content. Fetched public data is cached locally to reduce repeat downloads.

### 3. Local storage

The Extension uses `chrome.storage` to keep:

- a local cache of public buff data;
- your UI preferences (language, dark mode, layout).

All of it stays on your device and is never uploaded.

### 4. Permissions

| Permission | Purpose |
|---|---|
| `contextMenus` | Add a "Visualize" entry to the right-click menu when the user selects text |
| `sidePanel` | Display the Source / Target comparison interface in the browser side panel |
| `windows` | Allow the preview panel to open as a standalone window |
| `system.display` | Calculate the size and position of the standalone window on the current display |
| `storage` | Store UI preferences and a local cache of public data |
| `unlimitedStorage` | The public buff data set can exceed the default storage quota, so the cache needs more room |
| `alarms` | Check once a day for public data updates |
| `host_permissions: https://raw.githubusercontent.com/*` | Download public game buff data (JSON) from this endpoint. No user data is sent to it or any other server |
| `content_scripts: <all_urls>` | The core feature must read user-selected text on whatever page the user is on and render it locally; the set of sites cannot be predetermined |

### 5. Children's privacy

The Extension is not directed to children under 13 and collects no age-related information.

### 6. Changes

If this policy changes, we will update the effective date on this page and note it in the Extension's release notes.

### 7. Contact

For privacy questions, contact: 2581832218@qq.com
