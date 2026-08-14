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

### 1. 系统运行环境要求与依赖下载

YuLink 基于现代化 Chromium WebView2 与 .NET 体系构建。绝大多数 Windows 10 (2004+) 与 Windows 11 设备已预装以下环境；若系统缺失，可点击下方微软官方通道一键下载：

| 依赖组件 | 最低版本要求 | 微软官方在线安装包 | 微软官方离线独立安装包 | 说明 |
| :--- | :--- | :--- | :--- | :--- |
| **.NET Framework** | `4.8` 或 `4.8.1` | [🌐 在线安装](https://go.microsoft.com/fwlink/?LinkId=2085155) | [📦 离线安装包 (推荐)](https://go.microsoft.com/fwlink/?linkid=2088631) | 基础运行框架，提供 VSTO 与 C# 支持 |
| **Edge WebView2 Runtime** | Evergreen 最新版 | [🌐 在线安装 (推荐)](https://go.microsoft.com/fwlink/p/?LinkId=2124703) | [📦 x64 离线独立包](https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/038e5be3-91a2-4c14-b2eb-2fac728c8c2c/MicrosoftEdgeWebView2RuntimeInstallerX64.exe) | Chromium 极速网页渲染与 DOM 拦截内核 |
| **Office 宿主平台** | 2013 / 2016 / 2019 / 2021 / 365 或 WPS 演示 | 宿主本地安装即可 | - | 原生无缝集成 |

> [!TIP]
> 运行 **`一键安装.bat`** 时，脚本会自动诊断本机是否已就绪上述依赖。若检测到缺失，控制台会输出提示与直达下载链接。

---

### 2. 自动化工具与极简部署 (Tools & Installation)

YuLink 采用极致精简的“双脚本”设计，零门槛开箱即用：

| 脚本工具 | 功能说明 | 适用场景 |
| :--- | :--- | :--- |
| **`一键安装.bat`** 🚀 | **一键全自动部署**：多镜像智能诊断、静默补齐缺失依赖、脱机编译、写入证书信任与双宿主注册 | **首次安装 / 升级更新** |
| **`一键卸载.bat`** 🗑️ | **一键完全卸载**：清理所有注册表加载项、COM 项与 ClickOnce 信任残留 | **彻底移除插件** |

> [!TIP]
> **多通道镜像加速支持**：针对部分企业内网或网络受限环境，`一键安装.bat` 内置了**多通道镜像备用重试机制**（Microsoft CDN + 交付直连网络），当某个源超时时自动秒级切换至备用镜像，保障依赖下载 100% 成功。

#### 🚀 快速上手：
1. 下载本项目发布包或克隆代码仓库到本地。
2. 双击运行 **`一键安装.bat`**（全自动完成依赖自愈与插件注册）。
3. 打开 PowerPoint 或 WPS 演示，顶部将出现 **`YuLink`** 功能区标签，即刻开启使用！

#### 👨‍💻 开发者命令行构建：
```bash
git clone https://github.com/yosanji/YuLink.git
cd YuLink
powershell -ExecutionPolicy Bypass -File .\build_and_deploy.ps1
```

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

#### ⑤ 移动智能中枢三合一 (无线遥控 · 实物展台 · 屏幕投屏)
1. 在功能区点击 **`开启控制`** 并使用手机扫描 **`遥控二维码`**。
2. 手机端将进入多功能控制中枢：
   - 🎮 **智能遥控**：远程控制 PPT 翻页、视频播放与一键黑屏。
   - 📷 **无线实物展台 (零安装 App)**：
     - **📸 一键拍照投屏**：走下讲台拍摄学生作业，1秒全屏铺满 PPT 页面，大屏自动唤出画笔进行圈画批注！
     - **🎥 实时动态展台**：手机摄像头实时直播推流到 PPT 画面，适合理化生实验演示。
     - **🔄 90° 旋转**：支持大屏一键顺时针旋转，横竖版作业完美适应。
   - 📱 **手机屏幕镜像**：搭配开源无广告的 `ScreenStream`，一键将整台手机操作画面实时同屏到 PPT 页面。
   - ⏹️ **一键退出**：随时点击退出展台或投屏，PPT 瞬间恢复原幻灯片播放。


## 💬 交流与反馈 (Contact & Community)

如果您在教学课件制作、学术演讲演示中有任何改进建议，或遇到使用问题，欢迎交流探讨：

* 👤 **作者**：yosanji (鱼玄机)
* 💬 **微信**：`yosanji` (添加请备注 "YuLink" 或 "PPT插件")
* 🐛 **问题反馈**：[GitHub Issues](https://github.com/yosanji/YuLink/issues)

---

## 🤝 致谢与上游开源项目 (Acknowledgements & Credits)

本项目基于原作者 [@yuwenhui2020](https://github.com/yuwenhui2020) 开源的 [yuwenhui2020/PPT-Addin](https://github.com/yuwenhui2020/PPT-Addin) 进行了深度的架构重构与功能拓展。

特此向原作者在 PPT 网页内嵌领域的先驱探索致以诚挚敬意！YuLink 在此基础上重构了拟物悬浮批注系统、Localhost 透明反向代理、物理翻页笔直通及 Chrome 级离散缩放等现代演示体验。

---

## 📄 开源许可证 (License)

本项目采用 [MIT License](LICENSE) 开源协议。

Copyright (c) 2026 **yosanji (鱼玄机)**. All rights reserved.
