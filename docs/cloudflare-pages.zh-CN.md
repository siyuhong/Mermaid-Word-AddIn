# Cloudflare Pages 部署

**[English](cloudflare-pages.md) | 简体中文**

把 GitHub 仓库接入 Cloudflare Pages 后，每次推送生产分支都会自动执行 `npm run build`，并把 `dist/` 发布为 HTTPS 站点。用户侧载该站点生成的 `manifest.xml` 后，Word 就会从这个域名加载任务窗格。

## 1. Fork 本仓库

1. 打开 [siyuhong/Mermaid-Word-AddIn](https://github.com/siyuhong/Mermaid-Word-AddIn)。
2. 点击右上角 **Fork**，创建副本。

## 2. 创建 Cloudflare Pages 项目

1. 登录 [Cloudflare 控制台](https://dash.cloudflare.com/)。
2. 进入 **Workers & Pages → Create application → Pages → Connect to Git**。
3. 授权 Cloudflare 访问 GitHub，选择第 1 步创建的 fork 仓库。
4. 选择生产分支，例如 `main`。

## 3. 配置构建

按下表填写构建参数：

| 字段                   | 填写内容             |
| ---------------------- | -------------------- |
| Framework preset       | `None`               |
| Build command          | `npm run build`      |
| Build output directory | `dist`               |
| Root directory         | 留空，使用仓库根目录 |

构建环境变量：

| 变量名           | 值                                                                     |
| ---------------- | ---------------------------------------------------------------------- |
| `ADDIN_BASE_URL` | 该项目固定的生产地址，例如 `https://your-mermaid-word-demo.pages.dev/` |
| `NODE_VERSION`   | `22`（或本项目已验证的其他 Node LTS 版本）                             |

注意：

- 生产构建会把 `manifest.xml` 中所有 `https://localhost:3000` 以及对应的 `<AppDomain>` 替换为 `ADDIN_BASE_URL`。缺少该变量或地址不合法时，`npm run build` 会直接失败。
- `ADDIN_BASE_URL` 必须是 HTTPS 地址，且不能包含用户名、密码、查询串或锚点，也不要带 `/taskpane.html` 这类路径；结尾斜杠可有可无。
- 变量按环境区分。只在 Production 配置时，Preview（PR）部署会因为缺少 `ADDIN_BASE_URL` 而构建失败。若需要 PR 预览，请为 Preview 环境单独配置。

## 4. 部署

1. 点击 **Save and Deploy**，等待依赖安装、构建与上传完成。
2. 核对 Cloudflare 实际分配的生产域名，例如 `https://your-mermaid-word-demo.pages.dev`。
3. 如果域名与 `ADDIN_BASE_URL` 不一致，在 **Settings → Environment variables** 中修改 Production 环境变量后重新部署。

只改环境变量不会更新已经生成的清单，必须重新构建。绑定自定义域名后也一样：先改 `ADDIN_BASE_URL`，再重新构建，然后让用户重新下载清单。

## 5. 验证部署

### 5.1 浏览器检查

用实际生产域名访问以下路径：

| 路径                  | 预期                                                      |
| --------------------- | --------------------------------------------------------- |
| `/`                   | 显示安装说明和 `manifest.xml` 下载链接                    |
| `/manifest.xml`       | 下载 XML 文件，内容中所有地址指向生产站点而非 `localhost` |
| `/taskpane.html`      | 页面和脚本成功加载；Office.js 功能仍需在 Word 中验证      |
| `/commands.html`      | 页面能加载，内容为空白属正常现象                          |
| `/assets/icon-32.png` | 正常显示图标                                              |
| `/does-not-exist`     | 返回 `404.html`                                           |

若最后一项返回首页或 200，检查 `public/404.html` 是否被复制到 `dist/`，以及是否误配了 SPA 回退或重写规则。

### 5.2 在 Word 中验证

1. 从 `https://your-mermaid-word-demo.pages.dev/`（或 `/manifest.xml`）下载清单。
2. **Word 网页版**：打开可编辑文档，进入 **开始 → 加载项 → 更多设置 → 上传我的加载项**，选择清单。
3. **Windows 桌面版**：把清单放进共享文件夹，在 **信任中心 → 受信任的加载项目录** 中添加该文件夹的 UNC 路径，重启 Word 后从 **共享文件夹** 安装。
4. 在 Word 的 **插入** 选项卡中打开 **Mermaid Editor**。
5. 依次验证四条主路径：
   - 修改 Mermaid 代码，编辑器正常工作。
   - 实时预览随代码更新。
   - 点击 **Insert | Replace** 能将图表插入文档。
   - 选中已插入的图片，点击 **Edit Selected** 能恢复源码，修改后可再次替换。

## 6. 常见问题

- **构建报错 `Set ADDIN_BASE_URL ...`**：Production 环境变量未配置，或只配置到了 Preview。
- **构建报错 `ADDIN_BASE_URL must be a deployed HTTPS URL ...`**：地址不是 HTTPS，或带了查询串、锚点、账号密码，或仍指向 `localhost`。
- **Word 里任务窗格空白或报证书错误**：确认访问的是生产域名而不是 `localhost:3000`，并检查清单中的 URL 与当前域名是否一致。
- **重新部署后 Word 仍是旧版本**：`public/_headers` 已对全部资源设置 `no-cache`，但 Word 仍可能缓存清单。重新上传清单或重启 Word；必要时清理 WebView2 缓存。
- **任务窗格无法在 Word 网页版内嵌**：不要给站点添加 `X-Frame-Options`，或限制 `frame-ancestors` 的 CSP header。

## 官方参考

- [Cloudflare Pages Git 集成](https://developers.cloudflare.com/pages/get-started/git-integration/)
- [Cloudflare Pages 构建配置](https://developers.cloudflare.com/pages/configuration/build-configuration/)
- [Cloudflare Pages 响应头配置](https://developers.cloudflare.com/pages/configuration/headers/)
- [Office 网页版侧载加载项](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [Windows 共享目录侧载](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
