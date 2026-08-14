<p align="center">
  <h1 align="center">YuLink (鱼链)</h1>
  <p align="center">
    <strong>现代化幻灯片无缝网页嵌入与板书交互演示套件</strong>
    <br />
    <em>Next-Generation Seamless Web Embedding & Interactive Presentation Suite for PowerPoint & WPS</em>
  </p>
  <p align="center">
    <a href="#-what---这是什么">What 是什么</a> •
    <a href="#-why---为什么选择-yulink">Why 为什么选择</a> •
    <a href="#-how---如何安装与使用">How 如何使用</a> •
    <a href="README.md">English Documentation</a>
  </p>
  <p align="center">
    <a href="https://github.com/yosanji/YuLink/releases"><img src="https://img.shields.io/badge/Release-v1.0.0-blue.svg?style=flat-square&logo=github" alt="Version"></a>
    <img src="https://img.shields.io/badge/.NET_Framework-4.8.1-512BD4.svg?style=flat-square&logo=dotnet" alt=".NET">
    <img src="https://img.shields.io/badge/Chromium-WebView2-0078D7.svg?style=flat-square&logo=microsoftedge" alt="WebView2">
    <img src="https://img.shields.io/badge/Host-PowerPoint_%7C_WPS-D83B01.svg?style=flat-square&logo=microsoftpowerpoint" alt="Host">
    <img src="https://img.shields.io/badge/License-MIT-green.svg?style=flat-square" alt="License">
    <img src="https://img.shields.io/badge/Author-yosanji_(鱼玄机)-FF69B4.svg?style=flat-square" alt="Author">
  </p>
</p>

---

## 📌 WHAT - 这是什么？

> **YuLink** 是一款专为**教师、学术演讲者与演示工程师**打造的 Office / WPS 深度增强插件。

基于 **C# / VSTO** 技术与 **Microsoft Edge Chromium WebView2** 内核，它能够将现代 Web 网页（在线网站、本地 HTML5、3D 模型、音视频应用）**像素级原生嵌入到 PowerPoint 与 WPS 演示的放映画面中**。无需跳出 PPT，即可在全屏演讲过程中流畅进行实时交互操作与板书批注。

```
 ┌────────────────────────────────────────────────────────────────────────┐
 │                      PowerPoint / WPS 幻灯片放映视窗                   │
 │                                                                        │
 │   ┌────────────────────────────────────────────────────────────────┐   │
 │   │  🌐 YuLink 嵌入式浏览器 (Chromium WebView2)                     │   │
 │   │  • 3D 几何模型 (Three.js / GeoGebra)  • 课文互动跟读           │   │
 │   │  • 本地 Localhost Web 服务             • 流媒体网页视频         │   │
 │   │                                                                │   │
 │   │                             ┌──────────────────────────────┐   │   │
 │   │                             │ 🎨 拟物毛玻璃悬浮画笔工具     │   │   │
 │   │                             │ [🔴][🔵][🟢] 🖌️ 橡皮 撤销 扫码│   │   │
 │   │                             └──────────────────────────────┘   │   │
 │   └────────────────────────────────────────────────────────────────┘   │
 └────────────────────────────────────────────────────────────────────────┘
```

---

## 💡 WHY - 为什么选择 YuLink？

传统的课件制作与演讲汇报中，插入网页或视频往往面临**操作繁琐、画面割裂、无法批注、翻页失灵**等多重痛点。YuLink 专为真实讲台体验进行了全方位的技术革新。

### 1. 核心优势对比 (Why YuLink vs Traditional)

| 对比维度 | 传统超链接 / 切换浏览器 ❌ | 传统嵌入插件 ⚠️ | **YuLink (鱼链) ✅** |
| :--- | :--- | :--- | :--- |
| **放映连贯性** | 频繁 `Alt+Tab` 跳出，打断讲演节奏 | 需单独弹出窗口，遮挡背景 | **同屏原生无缝渲染，放映浑然一体** |
| **板书批注** | 切到浏览器后 PPT 画笔彻底失效 | 批注功能粗糙或无法交互 | **苹果拟物悬浮画笔，支持一键挂起交互** |
| **硬件翻页笔** | 鼠标点击网页后，物理翻页笔失灵 | 焦点被网页独占，按键卡壳 | **独家翻页事件直通，手持翻页笔永不卡壳** |
| **本地服务分享** | 手机无法访问电脑的 `localhost` | 只能分享静态外链 | **内置透明反向代理，手机扫码秒开本地资源** |
| **跨域与防嵌限制** | 报 `X-Frame-Options` 拒绝连接错误 | 页面直接白屏报错 | **底层请求拦截剥离限制头，全网网页秒开** |
| **画面缩放** | 浏览器缩放导致视口错位失真 | 缩放模糊拉伸 | **16 阶 Chrome 级离散等比缩放，像素对齐** |

---

### 2. 核心教学与演示场景

* 📐 **理科与数学 · 3D 立体几何与物理仿真**  
  原生加载 WebGL、Three.js、GeoGebra 等复杂模型。放映中直接拖拽旋转立体图形、动态展示受力与函数曲线。
* 📖 **文科与语言 · 课文互动跟读与多媒体点读**  
  嵌入具备语音测评、生词跟读与互动动画的 HTML5 课件，大屏即时互动。
* 🎨 **艺术与创意 · 动态数据大屏与生成式设计**  
  同屏流畅运行 Canvas 动效、ECharts / DataV 数据可视化与互动演示原型。
* 🎬 **全媒体集成 · 免跳转在线视频播放**  
  直接嵌入 Bilibili、YouTube 或内部视频流，省去外部播放器广告与跳转等待。

---

## 🛠️ HOW - 如何安装与使用？

### 1. 系统运行环境要求

* **操作系统**：Windows 10 / Windows 11 (64位)
* **支持宿主**：Microsoft PowerPoint (2013 / 2016 / 2019 / 2021 / Office 365) 或 WPS 演示 (WPP)
* **依赖环境**：
  - [.NET Framework 4.8.1](https://dotnet.microsoft.com/download/dotnet-framework/net481)
  - [Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/) (Win10/11 通常已自带)

---

### 2. 一键安装与卸载 (Installation)

#### 🚀 方式一：一键自动安装（推荐，小白友好）
1. 下载本项目发布包或克隆代码仓库到本地。
2. 双击运行目录下的 **`一键安装.bat`**。
3. 安装脚本将自动完成脱机编译、证书生成、ClickOnce 信任链注入及 Office 注册。
4. 打开 PowerPoint 或 WPS，菜单栏顶部将出现 **`YuLink`** 功能区！

#### 👨‍💻 方式二：开发者源码构建
```bash
git clone https://github.com/yosanji/YuLink.git
cd YuLink
powershell -ExecutionPolicy Bypass -File .\build_and_deploy.ps1
```

#### 🗑️ 一键卸载
若需从系统中完全清除，双击运行 **`一键卸载.bat`** 即可一键清理所有注册表项与信任缓存。

---

### 3. 功能详细操作手册 (User Guide)

```
  [ 插入网页 ] ──► [ 放映 F5 ] ──► [ 悬浮画笔 🖌️ ] ──► [ 扫码共享 📱 ] ──► [ 手机遥控 🎮 ]
```

#### ① 插入网页占位符
1. 打开 PowerPoint，点击上方功能区 **`YuLink`** 标签。
2. 点击 **`插入网页`**，页面将生成一个标有 `🌐 网页嵌入容器` 的矩形框。
3. 选中矩形框，在提示中输入您要展示的网址（如 `https://threejs.org` 或 `http://localhost:5173`）。

#### ② 幻灯片放映与交互
1. 按 **F5** 开启放映。翻到对应页面时，网页将秒级原生渲染。
2. 顶部 Safari 风格导航栏提供：**前进、后退、刷新、全屏、等比缩放** 控制。
3. **Chrome 级等比缩放**：点击 `+` / `-` 按钮可在 $25\% \sim 500\%$ 之间精确阶梯缩放，悬停即可查看当前精确百分比。

#### ③ 拟物悬浮画笔批注
1. 点击导航栏右侧的 **画笔** 图标，右下角将呼出磨砂玻璃悬浮圆盘。
2. **多色笔迹 & 撤销**：内置 4 种精选色系、3 种笔触粗细调节、橡皮擦与 30 步历史撤销。
3. **一键挂起交互**：点击当前选中的颜色即可快速暂停画笔，无缝恢复对底层网页的点击与滚动；再次点击恢复涂鸦。
4. **自适应展开**：圆盘拖动到屏幕左半区向右展开，拖到右半区向左展开，永不遮挡或超出屏幕边界。

#### ④ 本地 Localhost 扫码共享 (独家反代技术)
1. 展示本地开发网页（如 `http://localhost:3000`）时，点击导航栏的 **二维码** 图标。
2. 插件会自动激活内置反向代理，手机扫码后直接通过局域网拉取电脑本地的 HTML、CSS、JS 与图片等全部资产，轻松实现移动端同步展示。

#### ⑤ 局域网手机遥控器
1. 在功能区点击 **`开启控制`** 并扫描 **`遥控二维码`**。
2. 手机浏览器即可作为手持无线遥控面板，远程控制 PPT 的上一页、下一页翻页切换。

---

## 🤝 致谢与上游开源项目 (Acknowledgements & Credits)

本项目基于原作者 [@yuwenhui2020](https://github.com/yuwenhui2020) 开源的 [yuwenhui2020/PPT-Addin](https://github.com/yuwenhui2020/PPT-Addin) 进行了深度的架构重构与功能拓展。

特此向原作者在 PPT 网页内嵌领域的先驱探索致以诚挚敬意！YuLink 在此基础上重构了拟物悬浮批注系统、Localhost 透明反向代理、物理翻页笔直通及 Chrome 级离散缩放等现代演示体验。

---

## 📄 开源许可证 (License)

本项目采用 [MIT License](LICENSE) 开源协议。

Copyright (c) 2026 **yosanji (鱼玄机)**. All rights reserved.
