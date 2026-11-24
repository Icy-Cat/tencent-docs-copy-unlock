# 腾讯文档 · 解除复制限制（表格）

> 一个油猴脚本：解除腾讯文档**表格**的「禁止复制 / 禁止右键 / 禁止选择」限制，`⌘C` 直接把选中区域连同颜色、字体、边框、合并、对齐一起复制走。

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Tampermonkey%20%7C%20Violentmonkey-orange)

---

## 这是什么

腾讯文档可以给文档设「禁止复制」权限。开了之后，表格里 `⌘C` 会弹：

> 无法复制，管理员已禁止该权限

数据其实**已经在你的浏览器内存里了**，只是前端不让你写进剪贴板。这个脚本砍掉那层检查，然后自己从表格模型里读出选区，生成 `text/plain`（TSV）+ `text/html`（带样式的表格）写进剪贴板。

- 复制出来的内容**是带格式的**，粘到 Excel / 飞书 / 腾讯文档 / Google Sheets 都能保留底色和字型
- 表格是 canvas 渲染的，DOM 里没有文字，所以脚本读的是内部模型，不是选中文字

## 安装

1. 装 [Tampermonkey](https://www.tampermonkey.net/)（或 Violentmonkey）
2. 打开 [`tencent-docs-copy-unlock.user.js`](./tencent-docs-copy-unlock.user.js)，点 **Raw**
3. Tampermonkey 会弹出安装页 → 安装
4. **刷新**已经打开的腾讯文档表格页面（脚本必须抢在页面脚本之前运行，不刷新不生效）

## 用法

| 操作 | 效果 |
|---|---|
| `⌘C` / `Ctrl+C` | 复制当前选区（含格式） |
| `⌘⇧A` / `Ctrl+Shift+A` | 复制整张工作表（含格式） |
| `TDocsCopy.selection()` | 控制台调用，复制选区 |
| `TDocsCopy.sheet()` | 控制台调用，复制整表 |
| `TDocsCopy.dump()` | 只返回 `{text, html}`，不写剪贴板（调试用） |
| `TDocsCopy.last()` | 上一次复制的 payload |

复制成功会在页面底部弹一个蓝色小提示：`已复制 47 × 28（含格式）`。

## 会带 / 不会带的格式

**会带**（HTML 能表达的单元格级样式）

- 值本身（日期存日期值，长数字保持文本，不会变科学计数）
- 背景色、字体名、字号、加粗 / 斜体 / 下划线、字色
- 四条边框、自动换行、水平 / 垂直对齐
- 合并单元格（`rowspan` / `colspan`）、超链接

**不会带**（HTML 里没有这些概念，或腾讯会在粘贴时重排）

- **列宽、行高**
- 冻结窗格、筛选、条件格式
- 精确的数字格式（`0.00`、千分位、百分比、货币）——传的是**显示出来的文本**，粘过去由目标表格重新推断类型
- 批注、图片、图表、数据验证 / 下拉、工作表保护、分组、打印设置

> 如果目标文档也是腾讯文档，想连列宽行高一起搬，最省事的办法是从源表导出一份 `.xlsx` 再导入——不过「禁止复制」这类权限通常也会连带禁掉导出。

## 它到底做了什么

脚本分两层。

**第一层 · 通用解禁**（对任何页面都成立）

1. 包装 `EventTarget.prototype.addEventListener`，把所有 `copy / cut / contextmenu / selectstart / dragstart` 监听里的 `preventDefault` 换成空函数——事件照常派发（应用不会报错），但再也取消不掉浏览器默认动作
2. `oncopy / oncontextmenu / onselectstart` 这类 DOM0 属性一并放过
3. 注入 CSS：放开 `user-select`、隐藏水印层

**第二层 · 表格专用**

4. 在 `window` 捕获阶段接管 `⌘C`（比页面自己的监听更早），`stopImmediatePropagation` 掉它的键盘处理
5. 从内部模型读出选区：`SpreadsheetApp.view.canvas.spreadsheetView.getActiveSheetId()` / `normalSelectData._cpQ` / `getCellDataAtPosition(r, c)`
6. 拼 `text/plain`（TSV）和 `text/html`（`<table>` + inline style）
7. 写剪贴板：用一个隐藏容器 + `Range` 选区 + 自定义 `copy` 数据（**不能**直接 `execCommand('copy')`，没有 DOM 选区时它返回 `false`，剪贴板根本不会写）
8. 在自己的 `copy` 里也 `stopImmediatePropagation()`——腾讯的权限校验挂在 `copy` 监听里，不掐断它提示照样弹

## 踩过的坑（改代码前请先看）

| 坑 | 现象 | 处理 |
|---|---|---|
| 只拦 keydown 不够 | 不弹提示了但剪贴板是空的 | 权限校验在 `copy` 监听里，要在那里也 `stopImmediatePropagation()` |
| 表格容器本身是 `contenteditable` | `isEditable(target)` 把整个表格当成输入框，`⌘C` 直接放过 | 只把 `input/textarea` 当输入框，单元格编辑器单独判断 |
| `execCommand('copy')` 无选区 | 返回 `false`，看着「成功」其实没写 | 隐藏容器 + `Range` 选区 |
| 自己的补丁打到自己身上 | `preventDefault` 被置空，浏览器默认序列化覆盖掉我们 `setData` 的内容（粘出来是空格分隔、首尾多空行） | 给自己的监听打 `__dcopy` 标记，跳过补丁 |
| 注入时机 | `eval` 注入晚于页面脚本时，页面先把事件吃掉 | 必须 `@run-at document-start`，装完刷新页面 |

## 适用范围

**只支持腾讯文档的表格**（`docs.qq.com/sheet/...`）。

不支持，也别指望直接改 `@match` 就能用：

- 腾讯文档 · 文档（docx）/ 幻灯片 / 收集表——模型对象完全不同，且同样是 canvas 渲染，不读模型拿不到内容
- 飞书文档 / 石墨 / 语雀 / Notion 等——每家一套模型；对**文字在 DOM 里**的页面，第一层「通用解禁」单独拿出来就够用

## 已知限制

- 复制走的是系统剪贴板，如果这期间你在别处复制了东西，剪贴板会被覆盖（脚本无法避免）
- 大表（几万单元格）拼 HTML 会比较慢，`⌘C` 后可能卡一下
- 对齐枚举（`1 左 / 2 中 / 4 右`）是逆向经验值，`3`（常规）不输出，个别表格的对齐可能不准
- 只处理 `text/plain` + `text/html`，不生成 `application/x-rtf`，某些只认 RTF 的富文本编辑器会退化成纯文本

## 开发

没有构建，没有依赖，单文件。改完直接刷新页面（或重新装一次）即可。

测试时可以用 CDP 的 `Page.addScriptToEvaluateOnNewDocument({ source })` 把脚本注入到 `document-start`，等价于油猴的注入时机：

```js
await tab.cdp.send('Page.enable');
await tab.cdp.send('Page.addScriptToEvaluateOnNewDocument', { source: code });
await tab.evaluate(() => location.reload());
```

验证清单：

```js
// 页面已经加载脚本
typeof window.TDocsCopy           // 'object'

// 选区读得对不对
TDocsCopy.dump()                  // { text: "已处理标记\t制单日期...", html: "<html>..." }

// 剪贴板两种格式都在
const items = await navigator.clipboard.read();
items[0].types                    // ['text/plain', 'text/html']
```

## License

MIT
