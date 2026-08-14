<p align="center">
  <h1 align="center">YuLink (鱼链)</h1>
  <p align="center">
    <strong>Next-Generation Seamless Web Embedding & Interactive Presentation Suite for PowerPoint & WPS</strong>
    <br />
    <strong>为现代化教学与演示而生的无缝网页嵌入与板书交互增强套件</strong>
  </p>
  <p align="center">
    <a href="#-what---overview">What Overview</a> •
    <a href="#-why---why-yulink">Why YuLink</a> •
    <a href="#-how---installation--usage">How To Use</a> •
    <a href="README_zh-CN.md">简体中文完整文档</a>
  </p>
  <p align="center">
    <a href="https://github.com/yosanji/YuLink/releases"><img src="https://img.shields.io/badge/Release-v2.0.0-blue.svg?style=flat-square&logo=github" alt="Version"></a>
    <img src="https://img.shields.io/badge/.NET_Framework-4.8.1-512BD4.svg?style=flat-square&logo=dotnet" alt=".NET">
    <img src="https://img.shields.io/badge/Chromium-WebView2-0078D7.svg?style=flat-square&logo=microsoftedge" alt="WebView2">
    <img src="https://img.shields.io/badge/Host-PowerPoint_%7C_WPS-D83B01.svg?style=flat-square&logo=microsoftpowerpoint" alt="Host">
    <img src="https://img.shields.io/badge/License-MIT-green.svg?style=flat-square" alt="License">
    <img src="https://img.shields.io/badge/Author-yosanji_(鱼玄机)-FF69B4.svg?style=flat-square" alt="Author">
  </p>
</p>

---

## 📌 WHAT - Overview

> **YuLink** is a next-generation Office & WPS VSTO enhancement suite designed specifically for **educators, academic speakers, and presentation engineers**.

Powered by **C# / VSTO** and the **Microsoft Edge Chromium WebView2** engine, YuLink embeds modern Web applications (online sites, local HTML5, 3D WebGL models, audio/video streams) **pixel-perfectly into PowerPoint and WPS slideshows**. 

Presenters can interact with live web content and annotate directly on screen during fullscreen slideshows without disruptive window switching.

```
 ┌────────────────────────────────────────────────────────────────────────┐
 │                      PowerPoint / WPS Slideshow View                   │
 │                                                                        │
 │   ┌────────────────────────────────────────────────────────────────┐   │
 │   │  🌐 YuLink Embedded Browser (Chromium WebView2)                 │   │
 │   │  • 3D Geometry Models (Three.js / GeoGebra) • Read-Along Tools │   │
 │   │  • Localhost Dev Server                     • Streaming Media  │   │
 │   │                                                                │   │
 │   │                             ┌──────────────────────────────┐   │   │
 │   │                             │ 🎨 Floating Glassmorphism    │   │   │
 │   │                             │    Pen & Annotation Toolbar  │   │   │
 │   │                             │ [🔴][🔵][🟢] 🖌️ Undo  QR     │   │   │
 │   │                             └──────────────────────────────┘   │   │
 │   └────────────────────────────────────────────────────────────────┘   │
 └────────────────────────────────────────────────────────────────────────┘
```

---

## 💡 WHY - Why YuLink?

Traditional methods of presenting web content in slides suffer from **awkward window transitions, disabled annotation tools, broken clicker focus, and localhost isolation**. YuLink re-engineers the presentation experience from the ground up.

### 1. Feature Comparison (YuLink vs Traditional)

| Dimension | Traditional Alt+Tab / Hyperlink ❌ | Standard Embedded Plugins ⚠️ | **YuLink (鱼链) ✅** |
| :--- | :--- | :--- | :--- |
| **Presentation Flow** | Constant `Alt+Tab` breaks classroom rhythm | Popups block background content | **Seamless single-screen embedded rendering** |
| **On-Screen Ink Drawing** | PPT pen tools disappear outside slides | Crudely implemented or non-interactive | **Apple-style floating glass pen with 1-click pause** |
| **Wireless Presenter Clicker** | Mouse clicks on web steal focus; clicker dies | Keyboard focus trapped inside web view | **Exclusive IPC clicker pass-through; never locks up** |
| **Localhost Sharing** | Mobile devices cannot reach PC's `localhost` | Only shares static public links | **Built-in transparent reverse proxy; instant QR sync** |
| **Cross-Origin / CSP Blocks** | Blocked by `X-Frame-Options` errors | Blank screen errors | **Underlying request interception strips frame blocks** |
| **Viewport Zooming** | Browser zoom breaks layout and coordinates | Blurry pixel stretching | **16-step discrete Chrome-standard zoom scale** |

---

### 2. Core Educational & Professional Scenarios

* 📐 **STEM & Mathematics · Interactive 3D Geometric Simulation**  
  Natively run WebGL, Three.js, and GeoGebra models. Rotate 3D polyhedra, manipulate force vectors, and demonstrate dynamic calculus curves live during slideshows.
* 📖 **Language & Humanities · Interactive Read-Along & Phonetics**  
  Embed interactive reading courseware with text-to-speech assessment, dynamic phonetic cards, and stroke-order animations directly on the slide.
* 🎨 **Creative Design & Data Science · Live Visual Dashboards**  
  Present live ECharts / DataV dashboards, generative canvas animations, and interactive product prototypes with full fidelity.
* 🎬 **Seamless Rich Media · Zero-Jump Embedded Video Streaming**  
  Play Bilibili, YouTube, or internal streaming videos directly in-slide without browser tabs, ads, or buffer delays.

---

## 🛠️ HOW - Installation & Usage

### 1. System Requirements & Prerequisite Downloads

YuLink is built upon modern Chromium WebView2 and the .NET Framework. Most Windows 10 (2004+) and Windows 11 systems have these pre-installed. If your environment lacks them, download directly from official Microsoft links below:

| Component | Minimum Version | Microsoft Online Installer | Microsoft Standalone Offline Installer | Purpose |
| :--- | :--- | :--- | :--- | :--- |
| **.NET Framework** | `4.8` or `4.8.1` | [🌐 Web Installer](https://go.microsoft.com/fwlink/?LinkId=2085155) | [📦 Offline Installer (Recommended)](https://go.microsoft.com/fwlink/?linkid=2088631) | Base VSTO and C# presentation runtime |
| **Edge WebView2 Runtime** | Evergreen Latest | [🌐 Web Bootstrapper](https://go.microsoft.com/fwlink/p/?LinkId=2124703) | [📦 x64 Standalone Package](https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/038e5be3-91a2-4c14-b2eb-2fac728c8c2c/MicrosoftEdgeWebView2RuntimeInstallerX64.exe) | Chromium web engine & DOM interception |
| **Office Host** | 2013 / 2016 / 2019 / 2021 / 365 or WPS | Local installation | - | Native slideshow host |

> [!TIP]
> **`一键安装.bat`** automatically diagnoses these dependencies during installation and outputs direct download links if any are missing.

---

### 2. Streamlined Automation & One-Click Deployment (Tools & Installation)

YuLink adopts an ultra-clean "two-script" design with zero learning curve:

| Automation Tool | Description | Use Case |
| :--- | :--- | :--- |
| **`一键安装.bat`** 🚀 | **One-Click Auto Installer**: Multi-mirror auto-diagnosis, silent dependency auto-repair, offline compilation, ClickOnce trust injection, and dual-host registration. | **Initial Setup / Updates** |
| **`一键卸载.bat`** 🗑️ | **Clean Uninstaller**: Completely removes all registry add-in keys, COM registrations, and ClickOnce trust entries in 1 second. | **Full Uninstallation** |

> [!TIP]
> **Multi-Channel Mirror Acceleration**: For restricted networks or firewalled enterprise environments, `一键安装.bat` incorporates an automated multi-mirror fallback engine (Microsoft High-Speed CDN + Direct Delivery Network), seamlessly failing over to alternate mirrors if a connection times out.

#### 🚀 Quick Start:
1. Download the release package or clone this repository.
2. Double-click **`一键安装.bat`** (Handles auto-dependency repair & Office registration).
3. Launch PowerPoint or WPS — the **`YuLink`** tab is ready!

#### 👨‍💻 Developer Command-Line Build:
```bash
git clone https://github.com/yosanji/YuLink.git
cd YuLink
powershell -ExecutionPolicy Bypass -File .\build_and_deploy.ps1
```

---

### 3. Step-by-Step User Guide

```
  [ Insert Web ] ──► [ Slideshow F5 ] ──► [ Floating Pen 🖌️ ] ──► [ QR Share 📱 ] ──► [ Phone Remote 🎮 ]
```

#### ① Insert Web Page Placeholder
1. In PowerPoint, navigate to the **`YuLink`** ribbon tab.
2. Click **`插入网页` (Insert Web Page)**. A placeholder rectangle labeled `🌐 网页嵌入容器` will be created on the current slide.
3. Select the rectangle and enter the target URL (e.g. `https://threejs.org` or `http://localhost:5173`).

#### ② Fullscreen Slideshow & Interaction
1. Press **F5** to start your presentation. When navigating to the slide, the web page renders live in milliseconds.
2. Use the top Safari-style navigation bar for: **Back, Forward, Refresh, Fullscreen, and Discrete Zoom**.
3. **Chrome-Grade Zoom**: Click `+` / `-` to step through $25\% \sim 500\%$ discrete zoom levels with hover percentage tooltips.

#### ③ Glassmorphism Floating Annotation Pen
1. Click the **Pen** icon on the top navigation bar to summon the floating circular tool ball.
2. **Colors & Undo**: Features 4 curated color palettes, 3 stroke widths, an eraser, and a 30-step undo stack.
3. **One-Click Interaction Pause**: Click the active color dot again to temporarily suspend drawing mode and resume mouse clicking/scrolling on the underlying web page; click again to resume annotation.
4. **Adaptive Direction**: The tool tray automatically expands to the right when near the left edge, and to the left when near the right edge, never overflowing your display.

#### ④ Localhost QR Sharing (Reverse Proxy)
1. When presenting a local dev server (e.g. `http://localhost:3000`), click the **QR Code** icon in the navigation bar.
2. YuLink automatically activates its transparent reverse proxy. Mobile devices on the same LAN can scan and load all HTML, CSS, JS, and image assets seamlessly without requiring `--host` flags.

#### ⑤ Apple Minimalist Pure White Mobile Hub (Remote Controller · Wireless Visualizer)
1. Click **`开启控制` (Start Remote)** in the ribbon and scan the **`遥控二维码` (Remote QR)**.
2. The smartphone loads the clean, modern Apple white interface with haptic feedback:
   - 🎮 **Smart Remote Controller**: Flat minimalist layout with oversized Apple Blue "Next" control button, "Previous", and "Play/Pause" toggles with iOS-grade vibration feedback.
   - 📷 **Wireless Camera Visualizer (Zero App Installation, Scan & Cast)**:
     - **📸 Snap & Cast**: Walk around the classroom to snap student homework — loads onto the PPT slide in 1 second, with floating pen annotations ready to grade live!
     - **🔄 90° Screen Rotation**: Rotate portrait/landscape documents seamlessly.
     - **🎥 Live Video Stream**: Stream the phone's back camera directly onto the slide for live science experiments and demonstrations.
     - **💡 Torch**: Toggle camera flashlight for low-light environments.
     - **⏹️ One-Tap Exit**: Exit the visualizer anytime to smoothly resume the PowerPoint presentation.


## 💬 Contact & Community (交流与反馈)

Feel free to connect for ideas, teaching courseware sharing, feature suggestions, or feedback:

* 👤 **Author**: yosanji (鱼玄机)
* 💬 **WeChat**: `yosanji` (Please note "YuLink" when connecting)
* 🐛 **Bug Reports**: [GitHub Issues](https://github.com/yosanji/YuLink/issues)

---

## 🤝 Acknowledgements & Upstream Project (致谢与上游开源项目)

YuLink is built upon and inspired by the pioneering open-source work [PPT-Addin](https://github.com/yuwenhui2020/PPT-Addin) by [@yuwenhui2020](https://github.com/yuwenhui2020). 

We express our sincere gratitude to the original author for their foundation in Office web embedding. YuLink has deeply refactored the architecture, introducing glassmorphism floating annotations, transparent localhost proxying, hardware presenter remote pass-through, and Chrome-grade discrete zoom scaling.

---

## 📄 License (开源许可证)

This project is licensed under the [MIT License](LICENSE).

Copyright (c) 2026 **yosanji (鱼玄机)**. All rights reserved.
