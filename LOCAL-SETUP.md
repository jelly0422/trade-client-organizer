# 本地运行说明

完整源码目录：`C:\chen\AI+manju\trade-client-organizer`

## 方式一：双击启动（推荐）

1. 双击项目根目录的 `start-local.cmd`。
2. 保持命令窗口打开，在浏览器访问 http://127.0.0.1:4173 。
3. 用完后先导出客户数据，再在命令窗口按 Ctrl+C 停止服务。

如果提示 4173 端口被占用，请检查是否已经启动过本工具；已启动时直接访问上述地址即可。

## 方式二：直接打开网页

双击 `dist/index.html`，即可使用表格导入、字段映射、编辑及导出。无需联网，也无需启动 Node.js。

## 已备齐的依赖

| 依赖 | 版本 | 本地位置 |
| --- | --- | --- |
| SheetJS CE | 0.20.3 | `dist/vendor/xlsx.full.min.js` |
| Node.js Windows x64 LTS | 24.21.0 | `runtime/node-v24.21.0-win-x64/` |

Node.js 来自官方 https://nodejs.org/dist/v24.21.0/ ，下载压缩包已按同目录的 `SHASUMS256.txt` 校验 SHA-256。原始压缩包与校验文件保留在 `runtime/` 内。Node.js 官方发行包同时附带 npm，但本项目没有需要 npm 安装的依赖。

Excel 解析库、样式和脚本均随源码保存在本地；没有 CDN 请求、API Key 或后端数据库要求。

运行环境保留在项目内，不安装到系统，不修改系统 PATH。`runtime/` 已加入 `.gitignore`，避免把平台专用运行环境提交到 GitHub。

## 数据处理验证

在项目目录打开 PowerShell，执行：

```powershell
& .\runtime\node-v24.21.0-win-x64\node.exe --test .\tests\core.test.cjs
```

客户数据只在当前浏览器页面内存中处理，刷新或关闭之前请导出。
