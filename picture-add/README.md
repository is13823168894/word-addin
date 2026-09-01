# 插入图片（Word 加载项）

按页码、页面坐标或指定文字位置，向 Word 文档中插入图片。

## 文件说明

- `index.html`：任务窗格页面（含内联 CSS/JS）
- `manifest-picture-add.xml`：加载项清单文件

## 使用方式

1. 启动 HTTPS 服务器：

```powershell
npx http-server "e:/A001 重要软件/OFFICE加载项/word-addin" -p 8080 -S -C "$env:USERPROFILE/.office-addin-dev-certs/localhost.crt" -K "$env:USERPROFILE/.office-addin-dev-certs/localhost.key" --cors
```

2. 在 Word 中加载清单文件：
   - Word → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录
   - 添加目录 `e:/A001 重要软件/OFFICE加载项/word-addin/picture-add` 或共享目录
   - 重新启动 Word

3. 在"开始"选项卡中点击"插入图片"按钮。

## 功能

### 插入位置

- **按页码**：指定第 N 页插入（1=首页，-1=末页，负数从末页往前数）
- **按页面坐标**：指定图片在页面上的 X/Y 坐标（CM），提供「获取定位坐标」按钮快速获取参考位置
- **按文字后**：在指定字符串后面插入，可选择坐落方式：
  - 嵌入式紧随（与文字同行后）
  - 换行后左对齐 / 居中 / 右对齐

### 图片尺寸

- 原始尺寸（保持图片原始大小）
- 指定尺寸（可单独设宽或高，另一个自动按比例计算）

### 范围

- 当前文档（通过 Word JS API 实时操作）
- 多个 Word 文档（通过 JSZip 离线修改 XML）

## 注意事项

- 插入操作不可撤销，建议插入前保存文档
- 页面坐标系（X/Y）：X 为距左页边距距离，Y 为距段落上方距离（均为近似值）
- 多文档模式按页码插入时，基于 XML 分页符计数，可能与实际排版有偏差
- 「获取定位坐标」按钮依赖当前光标/选中的对象位置
