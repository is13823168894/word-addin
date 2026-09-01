## 自动换行 CSS 修改完成

已成功修改 `E:\A001 重要软件\OFFICE加载项\word-addin\digitalpanel\index.html`：

### 新增的 CSS 属性（第35行和第39-42行）：

```css
.np-table th {
    /* ... 其他样式 ... */
    word-break: break-all; 
    overflow-wrap: break-word; 
    white-space: normal;           /* 新增：允许正常换行 */
}

.np-table td {
    /* ... 其他样式 ... */
    word-break: break-all; 
    overflow-wrap: break-word; 
    white-space: normal;           /* 新增：允许正常换行 */
    /* 确保单元格能根据内容宽度自动调整，防止内容被强制压缩 */
    min-width: 0;
}
```

### 关键说明：

1. **`white-space: normal;`**: 关键属性，确保文本不会被锁定在 `nowrap` 模式下，允许浏览器根据需要自动换行。

2. **`min-width: 0;`**: 允许单元格收缩以适应内容宽度，防止长文本被强制压缩在一行内。

3. **`word-break: break-all;` + `overflow-wrap: break-word;`**: 已有属性，确保在单词边界或任意位置强制换行。

### 查看效果的建议：

由于 Office Web Viewer 可能会缓存旧版本，建议：
- **强制刷新**：在浏览器中按 `Ctrl + F5` 或 `Ctrl + Shift + R`
- **清除缓存**：清除浏览器缓存或使用无痕模式
- **重新加载加载项**：在 Word 中可能需要重新打开或刷新文档

现在当任何一列的文字超过约10个中文字符（或任何过长的文本）时，都应该能够自动换行显示，包括第一列的子表标签（如 `[40] 递延收益` 等）。

如果还有问题，可能涉及 Office Web Viewer 的特定渲染限制，建议尝试不同浏览器或检查 Word 文档的具体表现。