# 本地运行说明

项目目录：`C:\chen\AI+manju\trade-client-organizer`

## 已准备的本地版本

React 重构源码、npm 依赖和构建结果已保存在本地。双击项目根目录的 `start-local.cmd`，再访问 http://127.0.0.1:4173 。保持启动窗口打开；用完后先导出客户资料，再按 Ctrl+C 停止服务。

重构后网页使用 JavaScript 模块，需要本地 HTTP 服务；不再支持直接双击 `dist/index.html`。运行时的脚本、样式、图标及 Excel 解析库均在本地，无第三方 CDN、API Key 或后端数据库。

## 后续修改与构建

已安装 Node.js 24 或以上版本的电脑可运行：

```sh
npm ci
npm run dev
```

开发地址为 http://127.0.0.1:5173 。源码修改完成后运行：

```sh
npm test
npm run build
npm start
```

构建后的使用地址为 http://127.0.0.1:4173 。`dist/` 为生成文件，修改源码后必须重新构建；不要直接修改生成文件。

本机附带的便携 Node.js 位于 `runtime/node-v24.21.0-win-x64/`，无需全局安装。若系统找不到 npm，在项目目录的 PowerShell 使用：

```powershell
$env:PATH = (Join-Path (Get-Location) 'runtime/node-v24.21.0-win-x64') + ';' + $env:PATH
& .\runtime\node-v24.21.0-win-x64\node.exe .\runtime\node-v24.21.0-win-x64\node_modules\npm\bin\npm-cli.js run build
```

将末尾 `run build` 改为 `ci` 可重新安装依赖，改为 `run dev` 可启动开发服务器。首次安装依赖需要联网。

## 目录与数据

- `src/`：React / TypeScript 源码。
- `public/vendor/`：本地 SheetJS 解析库和许可证。
- `node_modules/`：已下载的前端依赖，Git 忽略。
- `dist/`：已生成的可运行网页，Git 忽略。
- `runtime/`：便携 Node.js 24.21.0，Git 忽略。
- `package-lock.json`：依赖版本锁文件，提交到 GitHub。

客户数据和字段设置仅保存在当前浏览器页面内存中。刷新或关闭前请导出；导出文件不是完整会话备份。
