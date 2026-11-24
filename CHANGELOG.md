# Changelog

## 1.1.0

- 修复：腾讯的权限校验挂在 `copy` 监听里，只在 keydown 阶段拦截时提示依旧会弹（补 `stopImmediatePropagation`）
- 修复：表格容器本身是 `contenteditable`，被误判成输入框导致 `⌘C` 直接放过
- 修复：`document.execCommand('copy')` 在没有 DOM 选区时返回 `false`，剪贴板并未写入（改为隐藏容器 + `Range` 选区）
- 修复：自身的 `preventDefault` 被自身的补丁置空，导致浏览器默认序列化覆盖了 `setData` 的数据
- 新增：`⌘⇧A` 复制整张工作表
- 新增：从 `fill.key` 解析背景色（`fill.fgColor` 常为空）
- 新增：`TDocsCopy.dump()` / `TDocsCopy.last()`

## 1.0.0

- 首个版本：解除复制 / 右键 / 选择限制，`⌘C` 复制选区（TSV + 带样式 HTML）
