# 批量删除图片（Word 加载项）

按范围、尺寸批量删除 Word 文档中的嵌入式图片、图形和文本框。

## 文件说明

- `index.html`：任务窗格页面（含内联 CSS/JS）
- `manifest-picture-del.xml`：加载项清单文件

## 使用方式

1. 将本目录部署到本地 Web 服务器，例如：

```powershell
npx http-server "e:/A001 重要软件/OFFICE加载项/word-addin" -p 8080 -S -C "$env:USERPROFILE/.office-addin-dev-certs/localhost.crt" -K "$env:USERPROFILE/.office-addin-dev-certs/localhost.key" --cors
```

2. 在 Word 中加载清单文件：
   - Word → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录
   - 添加本目录 `e:/A001 重要软件/OFFICE加载项/word-addin/picture-del` 或共享目录
   - 重新启动 Word

3. 在“开始”选项卡中点击“批量删除图片”按钮。

> 注意：若使用 nginx，请确保 `manifest-picture-del.xml` 中的 URL（`http://localhost/word-addin/picture-del/...`）已在 nginx 中正确映射到本目录。

## 功能

- 范围：当前文档（多个 Word 文档需逐个打开后执行）
- 类型：全部图片 / 指定尺寸的图片
- 可勾选删除对象：嵌入式图片、图形、文本框
- 支持从当前选中的图片自动获取尺寸

## 注意事项

- 删除操作不可撤销，建议先保存文档。
- 尺寸单位为 CM，与 Word 内部点值按 1 CM ≈ 28.35 pt 换算。
- “多个 Word 文档”选项当前需在 Word 中逐个打开文档后使用。
