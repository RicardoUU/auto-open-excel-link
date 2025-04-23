# Excel链接自动打开工具

这是一个基于Web的Excel链接解析工具，可以帮助用户快速打开Excel文件中的所有超链接。

## 功能特点

- 📄 上传并解析Excel文件 (.xlsx, .xls)
- 📑 支持多工作表(sheets)选择
- 🔗 提取并展示所有超链接
- 🚀 一键或批量打开所有链接
- ✨ 美观的用户界面

## 在线使用

访问 [https://ricardouu.github.io/auto-open-excel-link/](https://ricardouu.github.io/auto-open-excel-link/) 直接使用。

## 本地安装与运行

### 先决条件

- Node.js (版本 16.0.0 或更高)
- npm (通常随Node.js一起安装)

### 安装步骤

1. 克隆仓库
```bash
git clone https://github.com/ricardouu/auto-open-excel-link.git
cd auto-open-excel-link
```

2. 安装依赖
```bash
npm install
```

3. 启动开发服务器
```bash
npm run dev
```

4. 在浏览器中访问 `http://localhost:5173/auto-open-excel-link/`

### 构建生产版本

```bash
npm run build
```

构建后的文件将位于 `dist` 目录。

## 使用方法

1. 点击"选择文件"按钮或将Excel文件拖放到指定区域
2. 上传后，系统会自动解析Excel文件并读取所有工作表
3. 从下拉菜单中选择要查看的工作表
4. 系统会显示所选工作表中的所有超链接
5. 可以选择以下方式打开链接：
   - 点击"打开所有链接"按钮一次性打开所有链接
   - 使用复选框选择特定链接，然后点击"打开已选择的链接"
   - 单独点击表格或卡片中的链接
   - 使用"复制所有链接"功能

## 浏览器兼容性注意事项

大多数现代浏览器出于安全考虑会阻止网页一次性打开多个弹窗。如果您尝试一次性打开多个链接，可能需要：

1. 确认浏览器的弹窗拦截提示
2. 使用逐个打开功能
3. 手动Ctrl/Command+点击链接

## 技术栈

- React + TypeScript
- Vite
- XLSX库用于Excel文件解析

## 贡献

欢迎提交问题报告或贡献代码！

## 许可

MIT
