# IEEE Access 自动排版网页工具

这是一个基于 Flask 的网页端自动排版工具，目的是将文章内容与模板文档合并，生成符合模板样式的排版文章。

## 功能

- 上传文章文件 (.doc/.docx/.txt)
- 选择内置模板或上传自定义 .doc/.docx 模板
- 自动生成符合模板要求的排版文章 (.docx)
- 支持从文章文本中智能识别章节标题并应用模板样式
- 支持在网页端填写 SiliconFlow AI 配置，自动做章节切分与识别

## 运行方式

1. 安装依赖：

```bash
pip install -r requirements.txt
```

2. 启动服务：

```bash
python app.py
```

3. 在浏览器中访问：

```
http://127.0.0.1:5000
```

## 使用说明

- 默认可选择内置模板文件，当前项目已集成：
  - `HCIS_Template2026.docx`
  - `Access-Template.doc`
- 支持上传 `.doc` / `.docx` / `.txt` 文章与 `.doc` / `.docx` 模板，系统会尝试将 `.doc` 转换为 `.docx` 后再进行排版。
- 由于 `.doc` 转换依赖 Windows Word 自动化，建议在 Windows 环境中运行并安装 Word。
- 输出文件为 `.docx`，可在 Word 中打开并另存为 `.doc`。
- 若需要使用 AI 识别增强，可在网页端填写：
  - `AI API Key`
  - `AI 接口地址`
  - `AI 模型名`

也可以在项目根目录创建 `.env` 文件：

```ini
SILICONFLOW_API_URL=https://api.siliconflow.cn/v1/chat/completions
SILICONFLOW_API_KEY=你的密钥
SILICONFLOW_MODEL=Qwen/Qwen2.5-72B-Instruct
```

系统会自动调用 AI 进行章节切分和文本识别，提升中文和英文文档的识别准确度。
 
重要提示：请勿直接在聊天或公共场所粘贴您的 API 密钥。将其放入您机器上的 `.env` 文件中或将其设置为环境变量。助手无法使用粘贴到聊天中的密钥；请配置环境以便应用安全地读取它。
## 目录结构

- `app.py`: Flask 后端主程序
- `templates/index.html`: 前端页面
- `static/style.css`: 样式文件
- `requirements.txt`: Python 依赖
