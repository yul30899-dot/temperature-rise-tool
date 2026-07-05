# 温升数据分析工具 (Temperature Rise Data Organizer)

这是一个基于 React (Vite) + TailwindCSS + Electron 框架开发的桌面应用程序，主要用于批量整理温升测试数据，并一键生成带条件格式的 Excel 报告（及折线图）。

## 开发环境要求
- **Node.js**: 推荐 v18 或更高版本 (支持 npm 命令)

## 快速开始

1. **安装依赖**
   在项目根目录下打开命令行终端（CMD 或 PowerShell），执行以下命令：
   ```bash
   npm install
   ```
   *(这会自动下载所有的项目依赖库，并生成 `node_modules` 文件夹)*

2. **启动开发模式 (调试/热更新)**
   执行以下命令，同时启动本地前端服务并拉起 Electron 客户端窗口：
   ```bash
   npm run electron:dev
   ```
   *(在此模式下修改代码后，软件界面会实时更新，非常适合开发调试)*

## 打包成桌面软件 (.exe)

当你完成所有的代码修改，需要将项目重新打包成可以直接发给其他人使用的“绿色免安装版”时，只需执行：
```bash
npm run electron:build
```
打包完成后，您可以在项目根目录下的 `release/win-unpacked` 文件夹中找到编译好的软件。双击里面的 **`温升数据分析工具.exe`** 即可直接运行测试。

**提示**：若想把软件分享给同事，直接将 `win-unpacked` 整个文件夹打包压缩成 `.zip` 发送过去即可，对方解压后双击 `.exe` 就能完美运行。
