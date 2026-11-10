# Mermaid 图表 Word 加载项

[English](README.md) | **简体中文**

一个用于 Microsoft Word 的加载项，支持编写 Mermaid 代码、实时预览图表，并将图表以 PNG 图片插入文档。每张图片的替代文字中会保存对应的 Mermaid 源码，方便选中后再次编辑。

项目基于 React 18、Fluent UI v9 和 CodeMirror 6 构建，提供代码编辑与图表预览体验。

## 功能特性

- **交互式图表编辑器**：提供支持 Mermaid 语法高亮的代码编辑器
- **实时预览**：输入代码时自动更新图表预览
- **多种图表类型**：支持项目所依赖的 Mermaid 10.x 提供的图表语法，包括：
  - 流程图
  - 时序图
  - 类图
  - 状态图
  - 实体关系图
  - 甘特图
  - 饼图
  - 更多图表类型
- **语法高亮**：基于 CodeMirror 6，提供 Mermaid 专用语法高亮
- **自动补全**：补全 Mermaid 关键字、图表类型和语法
- **自动换行**：长行自动折行，便于阅读代码
- **预览导航**：支持放大、缩小、重置视图和拖动平移
- **主题选择**：支持 default、forest、dark 和 neutral 主题
- **编辑已有图表**：选中 Word 中已插入的图表，点击 **Edit Selected** 即可重新编辑
- **高质量输出**：以适合 Word 文档的高质量 PNG 图片插入图表

## 环境要求

- **Node.js**：使用最新的 LTS 版本。前往 [Node.js 官网](https://nodejs.org/)下载并安装，通过 `node -v` 和 `npm -v` 检查安装结果。
- **Git**：用于克隆仓库。
- **Word 桌面版**：以下命令面向桌面版 Word。使用项目自带的 Windows 调试配置时，建议安装较新版本的 Microsoft 365 Word，并使用 Edge Chromium/WebView2。不支持 Internet Explorer（Trident）和旧版 Edge（EdgeHTML）内核。

Mermaid 官网可能包含 10.x 尚不支持的新语法。使用新版文档中的示例时，请先确认项目安装的 Mermaid 版本。

## 安装与启动

1. 克隆仓库并进入项目目录：

   ```bash
   git clone https://github.com/siyuhong/Mermaid-Word-AddIn.git
   cd Mermaid-Word-AddIn
   ```

2. 安装依赖：

   ```bash
   npm install
   ```

3. 首次配置时，安装本地 HTTPS 开发证书，并在系统提示中信任该证书：

   ```bash
   npx office-addin-dev-certs install
   ```

4. 启动开发服务器，默认地址为 `https://localhost:3000`，并保持此终端运行：

   ```bash
   npm run dev-server
   ```

5. 在另一个终端中进入项目目录，将加载项侧载到 Word：

   ```bash
   npm start
   ```

使用结束后，执行 `npm stop` 移除开发加载项；在手动启动开发服务器的终端中按 `Ctrl+C` 停止服务器。

## 开发指南

### 使用 Office Add-ins Development Kit 运行加载项

1. **在 Visual Studio Code 中安装 Office Add-ins Development Kit**

   在 VS Code 扩展市场中搜索并安装 Office Add-ins Development Kit。

2. **预览加载项（F5）**

   选择 **Preview Your Office Add-in (F5)**，在快速选择菜单中选择 **Word Desktop (Edge Chromium)**。

   扩展会在调试前检查所需环境。检查完成后，Word 将启动并加载侧载的加载项。

3. **停止预览**

   测试结束后，选择 **Stop Previewing Your Office Add-in**，结束调试并移除侧载的加载项。手动启动的开发服务器需要单独停止。

### 项目命令

- `npm run build`：构建生产版本
- `npm run build:dev`：构建开发版本
- `npm run dev-server`：启动支持热更新的开发服务器
- `npm start`：将加载项侧载到 Word
- `npm stop`：从 Word 移除加载项
- `npm run validate`：校验 manifest 文件
- `npm run validate:production`：按生产要求校验构建后的 `dist/manifest.xml`
- `npm run lint`：检查代码风格
- `npm run lint:fix`：自动修复代码风格问题
- `npm run prettier`：格式化源代码
- `npm run watch`：监听源文件变化并重新构建

### 生产部署

`npm run build` 将生产构建产物写入 `dist/`，不会自动发布。生产构建要求设置环境变量 `ADDIN_BASE_URL`，不再使用占位地址，也无需把生产域名写进源 manifest。

```powershell
$env:ADDIN_BASE_URL = "https://你的项目名.pages.dev/"
npm run build
```

构建自动生成线上页面、图标、帮助链接和允许域名对应的 `dist/manifest.xml`；根目录的 `manifest.xml` 仍用于本地开发。将 `dist/` 的内容部署后，运行 `npm run validate:production`，并把站点首页或 `dist/manifest.xml` 交给试用用户。

完整步骤见 [Cloudflare Pages 部署与侧载指南](docs/cloudflare-pages.zh-CN.md)，包括 Git 自动部署、手动上传和用户安装说明。侧载属于微软支持的开发测试方式，不是商店发布或企业集中部署的替代渠道。

## 使用方法

1. 打开 Word，在“插入”选项卡中找到 **Mermaid Editor** 按钮。
2. 点击按钮打开任务窗格。
3. 在编辑器中编写 Mermaid 代码，或直接使用默认示例。
4. 在预览区域查看渲染结果。
5. 根据需要在下拉框中选择主题。
6. 等待预览完成，将 Word 光标放在目标位置，点击 **Insert | Replace** 插入图表。如果 Word 当前选中了文字或图片，该选区会被替换。
7. 如需编辑已有图表：
   - 在 Word 中选中图表图片。
   - 点击任务窗格中的 **Edit Selected**。
   - 图表源码将载入编辑器。
   - 修改代码，等待预览更新完成。
   - 保持原图选中，或重新选中原图，再点击 **Insert | Replace** 替换图片。

代码或主题变更尚未完成渲染时，插入按钮会保持禁用。**Edit Selected** 适用于由本加载项插入、且替代文字中仍保留 Mermaid 标题和源码的嵌入型图片；如需再次编辑，请保留这些信息。替换始终作用于 Word 当前选区，不会自动绑定到之前读取源码的图片。

如果 Mermaid 初始化指令或 YAML 配置中显式指定了主题，该配置会优先于主题下拉框。

## 项目结构

- `/src/taskpane/`
  - `index.jsx`：React 入口与 Office.js 初始化
  - `taskpane.js`：Word JavaScript API 集成
  - `components/MermaidEditor.jsx`：主编辑器组件
  - `components/enhancedMermaidLanguage.js`：CodeMirror 的 Mermaid 语言支持增强
  - `utils/diagramUtils.js`：图表类型检测与工具函数
  - `utils/mermaidRenderer.js`：Mermaid 渲染队列与导出配置
- `/assets/`：加载项图标与图片
- `/src/commands/`：加载项命令回调与宿主页面
- `/dist/`：构建生成目录，不应手动编辑
- `manifest.xml`：加载项清单配置
- `webpack.config.js`：Webpack 打包配置

## 主要技术

- **React 18**：用户界面框架
- **Fluent UI v9**：微软设计系统
- **CodeMirror 6**：支持语法高亮的代码编辑器
- **Mermaid.js**：图表渲染引擎
- **Office.js**：Word JavaScript API
- **Webpack**：模块打包工具

## 常见问题排查

如果加载项无法正常运行，可以尝试：

- 关闭所有已打开的 Word 窗口。
- 通过 **Stop Previewing Your Office Add-in** 或 `npm stop` 停止调试；手动启动的开发服务器请在对应终端中按 `Ctrl+C` 停止。
- 如果遇到 HTTPS 信任错误，运行 `npx office-addin-dev-certs verify` 检查证书，必要时重新安装。
- 使用本地侧载的加载项时，确认开发服务器仍在 `https://localhost:3000` 运行。
- 如果问题仍然存在，清理 Office 缓存。

更多帮助请参阅 [Office 加载项开发错误排查](https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-development-errors)，或[向本项目反馈问题](https://github.com/siyuhong/Mermaid-Word-AddIn/issues)。

## 参考资源

- [Office 加载项文档](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Word JavaScript API 参考](https://learn.microsoft.com/javascript/api/word)
- [Mermaid 文档](https://mermaid.js.org/)
- [在 Office 网页版中侧载加载项](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

## 许可证

MIT License，详见 [LICENSE](LICENSE)。

Copyright (c) 2024 Microsoft Corporation. All rights reserved.
