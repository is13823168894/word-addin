# 词汇增强助手 - Word 加载项

Word 加载项，用于词汇统计高亮、查找定位、批量替换。托管在 GitHub Pages，无需本地服务器。

## 在线安装（推荐）

1. 在 GitHub 上创建仓库 `word-vocabulary-addin`，推送代码
2. 开启 GitHub Pages：仓库 → Settings → Pages → Source 选 `main` 分支 → Save
3. 等待部署完成（约1分钟），访问地址：`https://jeffersongg.github.io/word-vocabulary-addin/taskpane.html`
4. 在 Word 中侧边加载 `manifest.xml`：
   - 打开 Word → 插入 → 获取加载项 → 管理我的加载项 → 上载我的加载项 → 选择 `manifest.xml`

## 更新插件

修改 `taskpane.html` → 推送到 GitHub → Word 侧边栏刷新即可

## 功能

- **统计并高亮**：扫描文档中所有关键词，黄色高亮显示
- **关键词标签**：输入的关键词自动生成可点击标签，点击即可查找定位
- **批量替换**：按行对应关系一次性替换所有词汇
- **清空全部**：一键清除所有输入和高亮

## 文件说明

| 文件 | 说明 |
|------|------|
| `manifest.xml` | 加载项清单配置 |
| `taskpane.html` | 主界面（所有代码内联） |
| `README.md` | 说明文档 |

## 注意事项

- GitHub Pages 必须开启 HTTPS（默认已开启）
- 图标文件 `icon-16.png`、`icon-32.png`、`icon-80.png` 可选，没有不影响功能
- 【备份】开头的文件是备份，不要修改或删除
