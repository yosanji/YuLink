# YuLink (鱼链) - 现代幻灯片无缝网页嵌入与教学交互演示套件

<p align="center">
  <img src="https://img.shields.io/badge/YuLink-v1.0.0-blue.svg?style=for-the-badge&logo=github" alt="Version">
  <img src="https://img.shields.io/badge/.NET_Framework-4.8.1-512BD4.svg?style=for-the-badge&logo=dotnet" alt=".NET">
  <img src="https://img.shields.io/badge/Chromium-WebView2-0078D7.svg?style=for-the-badge&logo=microsoftedge" alt="WebView2">
  <img src="https://img.shields.io/badge/Host-PowerPoint_%7C_WPS-D83B01.svg?style=for-the-badge&logo=microsoftpowerpoint" alt="Host">
  <img src="https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge" alt="License">
  <img src="https://img.shields.io/badge/Author-yosanji_(鱼玄机)-FF69B4.svg?style=for-the-badge" alt="Author">
</p>

> **让网页的无限交互力，原生融入每一页幻灯片。**  
> **YuLink** 是一款专为**教育工作者、学术演讲者与演示工程师**打造的次世代 Office 增强套件。基于 **C# / VSTO** 与 **Microsoft Edge Chromium WebView2** 内核，它消除了 PPT 与浏览器之间割裂的窗口跳转，将现代 Web 技术的生动表达能力无缝注入 PowerPoint 与 WPS 演示之中。

---

## 💡 写在前面 · 为什么开发 YuLink

作为一名一线教育工作者，在日常教学中我发掘了 **HTML5 现代网页在课堂教学中不可替代的奇效**。无论是生动的动态物理仿真、三维几何模型，还是沉浸式的课文互动跟读，网页端都展现出了远超传统静态 PPT 的表现力。

然而，传统的教学演示往往面临三大割裂痛点：
1. **频繁跳转破坏节奏**：在 PPT 与外部浏览器之间频繁 `Alt + Tab` 切换，打断教学连贯性。
2. **切出后无法板书批注**：一旦切到外部网页，PPT 自带的画笔瞬间失效，无法在关键知识点上随手圈画。
3. **本地开发模型难以共享**：许多先进的教学模型运行在本地（如 Localhost、本地 WebGL 课件），缺乏免配置的同屏演示与移动端同步方案。

**YuLink 应运而生。** 本项目在开源社区优秀探索的基础上进行了深度重构与功能拓展，旨在**让幻灯片演播回归纯粹、流畅与一体化**。

---

## 🎯 教学与演示应用场景

```
                               ┌───────────────────────────────────────────────┐
                               │                 YuLink 演示体系                │
                               └──────────────────────┬────────────────────────┘
                                                      │
         ┌────────────────────────┬───────────────────┴────────────────┬────────────────────────┐
         ▼                        ▼                                    ▼                        ▼
 📐 理科与数学教学        📖 文科与语言学习                    🎨 创意与设计教学        🎬 全媒体融合播放
  • 3D WebGL (Three.js)    • 实时课文互动跟读                   • 生成式动态艺术         • 免跳转内嵌网页视频
  • 交互式 GeoGebra 模型   • 拼音生词点读卡片                   • ECharts / DataV 大屏   • 零等待流畅同屏播放
```

---

## ✨ 核心功能与亮点

- 🎨 **苹果拟物悬浮画笔**：自适应方向展开（靠左向右展、靠右向左展，永不溢出），支持一键挂起交互（选择同色即暂停涂鸦接管网页操作）。
- 📱 **Localhost 透明反向代理**：针对本地开发服务（如 `http://localhost:3000` / Vite / Webpack / Python Server），内置轻量级透明反代。手机扫码直接拉取电脑本地网页、CSS、JS 与图片等全部资产，无需开启 `--host`。
- 🎯 **物理无线翻页笔直通**：专为讲台实战设计。即使鼠标点击网页抢占了 Windows 焦点，手里的**物理无线翻页笔**（PageDown / PageUp / 空格 / 方向键）依然能丝滑驱动 PPT 翻页。
- 🔍 **Chrome 级离散等比缩放**：提供 $25\%$ 至 $500\%$ 的 16 阶 Chrome 标准百分比缩放，彻底消除局部错位拉伸。
- 🛡️ **解除 X-Frame-Options 安全防嵌**：底层网络请求拦截，自动剥离安全防嵌套响应头，拒绝页面白屏。
- 🌐 **双平台深度兼容**：原生适配 **Microsoft PowerPoint (2013-365)** 与 **WPS 演示 (WPP)**。

---

## 🚀 快速安装

### 方式一：一键自动安装（推荐）
1. 下载本项目发布包或克隆仓库至本地。
2. 双击运行目录下的 **`一键安装.bat`**。
3. 脚本自动完成脱机编译、自签名证书生成、信任链写入与 Office 注册。
4. 打开 PowerPoint 或 WPS，顶部将出现 **`YuLink`** 功能区标签。

### 方式二：开发者源码构建
```bash
git clone https://github.com/yosanji/YuLink.git
cd YuLink
powershell -ExecutionPolicy Bypass -File .\build_and_deploy.ps1
```

---

## 📖 交互操作指南

1. **插入网页**：点击 PPT 菜单栏 **`YuLink`** -> **`插入网页`**，在生成的占位符中输入目标网址（如 `http://localhost:5173` 或在线网页）。
2. **放映演示**：按下 **F5** 开启放映，网页秒级渲染就绪。顶部控制栏提供前进、后退、刷新、全屏与缩放功能。
3. **画笔标注**：点击顶部 **画笔** 图标唤出悬浮工具球，随时圈画重点。
4. **扫码共享**：点击 **二维码** 图标，听众或学生扫码即可同步查阅展示内容。
5. **局域网遥控**：点击 **`开启控制`** 并扫描 **`遥控二维码`**，手机化身掌上无线遥控器。

---

## 🗑️ 一键卸载

如需从系统中彻底移除插件，双击运行 **`一键卸载.bat`** 即可完全清理所有注册表项与信任证书缓存。

---

## 🤝 致谢与开源说明 (Acknowledgements)

本项目在开源社区先驱探索的基础上进行了深度重构与架构升级。特别致敬为 Office 开发者生态做出杰出贡献的开源项目与工程师群体。

---

## 📄 开源许可证 (License)

本项目采用 [MIT License](LICENSE) 开源协议。

Copyright (c) 2026 **yosanji (鱼玄机)**. All rights reserved.
