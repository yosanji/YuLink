<p align="center">
  <h1 align="center">YuLink (鱼链)</h1>
  <p align="center">
    <strong>Next-Generation Seamless Web Embedding & Interactive Presentation Suite for PowerPoint & WPS</strong>
    <br />
    <strong>为现代化教学与演示而生的无缝网页嵌入与板书交互增强套件</strong>
  </p>
  <p align="center">
    <a href="#english">English</a> •
    <a href="#-中文说明">中文说明</a> •
    <a href="README_zh-CN.md">简体中文文档</a>
  </p>
  <p align="center">
    <a href="https://github.com/yosanji/YuLink/releases"><img src="https://img.shields.io/badge/Release-v1.0.0-blue.svg?style=for-the-badge&logo=github" alt="Version"></a>
    <img src="https://img.shields.io/badge/.NET_Framework-4.8.1-512BD4.svg?style=for-the-badge&logo=dotnet" alt=".NET">
    <img src="https://img.shields.io/badge/Chromium-WebView2-0078D7.svg?style=for-the-badge&logo=microsoftedge" alt="WebView2">
    <img src="https://img.shields.io/badge/Host-PowerPoint_%7C_WPS-D83B01.svg?style=for-the-badge&logo=microsoftpowerpoint" alt="Host">
    <img src="https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge" alt="License">
    <img src="https://img.shields.io/badge/Author-yosanji_(鱼玄机)-FF69B4.svg?style=for-the-badge" alt="Author">
  </p>
</p>

---

<a name="english"></a>
## 🇬🇧 English

### 💡 Why YuLink?

As an frontline educator, I discovered the extraordinary pedagogical power of **modern interactive web pages** in classroom teaching. From dynamic 3D geometric models to real-time audio-visual read-along tools, web technologies bring unprecedented engagement that traditional static PowerPoint slides simply cannot match.

However, traditional presentation workflows suffer from significant friction:
- **Disruptive Alt-Tab Switching**: Jumping between PPT and external browsers breaks the narrative flow and disrupts student focus.
- **Lost Annotation Capabilities**: The moment you switch out of PPT, built-in presentation drawing tools disappear, making real-time on-page explanation nearly impossible.
- **Localhost Isolation**: Modern interactive apps running locally (e.g. `localhost:3000`, local 3D models) lack a zero-config, single-screen embedding workflow.

**YuLink was built to eliminate these boundaries.** Built on **C# / VSTO** and **Microsoft Edge Chromium WebView2**, YuLink fuses the infinite interactivity of the Web directly into PowerPoint and WPS slideshows.

---

### ✨ Key Features

| Feature | Description |
| :--- | :--- |
| 🎨 **Skeuomorphic Floating Pen** | iOS-inspired glassmorphism floating tool ball with adaptive auto-expanding direction (left/right aware) and 30-step undo history. |
| 📱 **Transparent Localhost Proxy** | Built-in zero-config reverse proxy for local dev servers (`localhost:3000`, Vite, LiveServer). Mobile devices can scan and load local pages and assets instantly. |
| 🎯 **Presenter Clicker Pass-Through** | Hardware wireless clicker keys (`PageDown`, `PageUp`, `Space`) pass straight through to PPT slide navigation even when the web page has mouse focus. |
| 🔍 **Discrete Proportional Zoom** | 16-step Chrome-standard discrete zoom scale ($25\%$ to $500\%$) with pixel-perfect viewport alignment and hover percentage tooltips. |
| 🛡️ **Bypass Frame Restrictions** | Deep request interception strips `X-Frame-Options` and `CSP` headers to guarantee cross-origin iframe-free rendering. |
| 🌐 **Dual Host Support** | Native integration with both **Microsoft PowerPoint (2013-365)** and **Kingsoft WPS Presentation (WPP)**. |

---

### 🎯 Educational & Presentation Scenarios

```
                               ┌───────────────────────────────────────────────┐
                               │              YuLink Presentation             │
                               └──────────────────────┬────────────────────────┘
                                                      │
         ┌────────────────────────┬───────────────────┴────────────────┬────────────────────────┐
         ▼                        ▼                                    ▼                        ▼
 📐 STEM & Geometry       📖 Language & Reading                🎨 Creative & Art        🎬 Seamless Media
  • 3D WebGL (Three.js)    • Real-time Read-Along               • Generative Art         • Embedded Web Videos
  • Interactive GeoGebra   • Phonetic Interactive Cards         • ECharts / DataV        • Zero-Jump Stream
```

---

### ⚡ Quick Start

#### Method 1: One-Click Installer (Recommended)
1. Download or clone this repository to your computer.
2. Double-click **`一键安装.bat`** (One-Click Install).
3. The script will automatically compile, generate local certificates, configure ClickOnce trust, and register the add-in.
4. Launch PowerPoint or WPS — the **`YuLink`** tab will appear on your ribbon.

#### Method 2: Developer Source Build
```bash
git clone https://github.com/yosanji/YuLink.git
cd YuLink
powershell -ExecutionPolicy Bypass -File .\build_and_deploy.ps1
```

---

### 🗑️ Uninstallation
Double-click **`一键卸载.bat`** to completely remove all registry entries, trust certificates, and cache files.

---

<br />

---

<a name="-中文说明"></a>
## 🇨🇳 中文说明

### 💡 写在前面 · 为什么开发 YuLink

作为一名一线教育工作者，在日常教学中我发掘了 **HTML5 现代网页在课堂教学中不可替代的奇效**。无论是生动的动态物理仿真、三维几何模型，还是沉浸式的课文互动跟读，网页端都展现出了远超传统静态 PPT 的表现力。

然而，传统的教学演示往往面临三大割裂痛点：
1. **频繁跳转破坏节奏**：在 PPT 与外部浏览器之间频繁 `Alt + Tab` 切换，打断教学连贯性。
2. **切出后无法板书批注**：一旦切到外部网页，PPT 自带的画笔瞬间失效，无法在关键知识点上随手圈画。
3. **本地开发模型难以共享**：许多先进的教学模型运行在本地（如 Localhost、本地 WebGL 课件），缺乏免配置的同屏演示与移动端同步方案。

**YuLink 应运而生。** 本项目基于 **C# / VSTO** 与 **Microsoft Edge Chromium WebView2** 内核，旨在**让网页的无限交互力，原生融入每一页幻灯片**。

---

### ✨ 核心功能与亮点

- 🎨 **苹果拟物悬浮画笔**：自适应方向展开（靠左向右展、靠右向左展，永不溢出），支持一键挂起交互（选择同色即暂停涂鸦接管网页操作）。
- 📱 **Localhost 透明反向代理**：针对本地开发服务（如 `http://localhost:3000` / Vite / Webpack / Python Server），内置轻量级透明反代。手机扫码直接拉取电脑本地网页、CSS、JS 与图片等全部资产，无需开启 `--host`。
- 🎯 **物理无线翻页笔直通**：专为讲台实战设计。即使鼠标点击网页抢占了 Windows 焦点，手里的**物理无线翻页笔**（PageDown / PageUp / 空格 / 方向键）依然能丝滑驱动 PPT 翻页。
- 🔍 **Chrome 级离散等比缩放**：提供 $25\%$ 至 $500\%$ 的 16 阶 Chrome 标准百分比缩放，彻底消除局部错位拉伸。
- 🛡️ **解除 X-Frame-Options 安全防嵌**：底层网络请求拦截，自动剥离安全防嵌套响应头，拒绝页面白屏。
- 🌐 **双平台深度兼容**：原生适配 **Microsoft PowerPoint (2013-365)** 与 **WPS 演示 (WPP)**。

---

### 🎯 教学与演示应用场景

* 📐 **理科与数学教学 · 3D 模型深度交互**：原生嵌入 Three.js、GeoGebra 等三维几何与物理仿真，放映中直接拖拽旋转立体模型。
* 📖 **文科与语言学习 · 课文互动跟读**：嵌入具备语音朗读、生词点读与互动测评的课文网页。
* 🎨 **艺术与创意设计 · 生成式交互演示**：展示 Canvas 动态大屏（DataV / ECharts）与生成式艺术。
* 🎬 **全媒体融合 · 无跳转流媒体播放**：免去超链接跳出与广告等待，网页视频同屏原生播放。

---

### 🚀 快速安装

#### 方式一：一键自动安装（推荐）
1. 下载本项目发布包或克隆仓库至本地。
2. 双击运行目录下的 **`一键安装.bat`**。
3. 脚本自动完成脱机编译、自签名证书生成、信任链写入与 Office 注册。
4. 打开 PowerPoint 或 WPS，顶部将出现 **`YuLink`** 功能区标签。

#### 方式二：开发者源码构建
```bash
git clone https://github.com/yosanji/YuLink.git
cd YuLink
powershell -ExecutionPolicy Bypass -File .\build_and_deploy.ps1
```

---

### 🗑️ 一键卸载
如需从系统中彻底移除插件，双击运行 **`一键卸载.bat`** 即可完全清理所有注册表项与信任证书缓存。

---

## 🤝 致谢与开源说明 (Acknowledgements)

本项目在开源社区先驱探索的基础上进行了深度重构与架构升级。特别致敬为 Office 开发者生态做出杰出贡献的开源项目与工程师群体。

---

## 📄 开源许可证 (License)

本项目采用 [MIT License](LICENSE) 开源协议。

Copyright (c) 2026 **yosanji (鱼玄机)**. All rights reserved.
