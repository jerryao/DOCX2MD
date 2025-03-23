# DOCX 批量转换 Markdown 工具

这是一个基于 [MarkItDown](https://github.com/microsoft/markitdown) 库的 Web 应用，用于批量将 DOCX 文档转换为 Markdown 格式。

## 功能特点

- 同时上传多个 DOCX 文件进行批量转换
- 保持与原始 DOCX 文件相同的文件名
- 两种导出方式：
  - 下载 ZIP 压缩包（默认）
  - 直接保存到指定本地目录
- 支持中文文档转换
- 保留文档结构（标题、表格、列表等）
- 直观的用户界面，包含文件预览和转换状态
- 异步处理，避免页面卡顿

## 安装使用

### 方法一：直接安装

1. 确保安装了 Python 3.8 或更高版本

2. 克隆或下载此仓库

3. 安装依赖项：
   ```
   pip install -r requirements.txt
   ```

4. 运行应用：
   ```
   python app.py
   ```

5. 打开浏览器访问：`http://localhost:5000`

### 方法二：使用 Docker

1. 确保已安装 Docker

2. 构建 Docker 镜像：
   ```
   docker build -t docx-to-markdown .
   ```

3. 运行容器：
   ```
   docker run -p 5000:5000 docx-to-markdown
   ```

4. 打开浏览器访问：`http://localhost:5000`

## 使用说明

1. 在网页界面中选择要转换的 DOCX 文件（可多选）
2. 选择导出方式：
   - **ZIP 压缩包**：转换完成后自动下载包含所有 Markdown 文件的压缩包
   - **指定目录**：输入本地目录路径，转换后的文件将直接保存到该目录
3. 点击"上传并转换"按钮
4. 等待转换完成

## 导出目录说明

当选择"保存到指定目录"选项时：

- 需要提供一个有效的本地文件路径，例如：`C:\Exports\Markdown` 或 `/home/user/exports`
- 如果目录不存在，系统会尝试自动创建
- 每个 Markdown 文件将保持与原 DOCX 文件相同的文件名
- 如果目录中已存在同名文件，将被覆盖

## 生产环境部署

对于生产环境，建议：

1. 设置强密钥：
   ```
   export SECRET_KEY="your-strong-secret-key"
   ```

2. 使用 Gunicorn 或 uWSGI 作为 WSGI 服务器

3. 配置反向代理 (Nginx/Apache)

4. 增加文件上传限制 (如需要)

## 技术栈

- 后端：Flask + MarkItDown
- 前端：Bootstrap 5 + JavaScript
- 容器化：Docker

## 注意事项

- 上传文件大小限制为 50MB
- 目前仅支持 .docx 格式文件
- 转换文件会临时保存，转换完成后自动删除
- 下载的 ZIP 文件会在服务器上保存 24 小时后自动删除

## 许可证

MIT

## 致谢

- [MarkItDown](https://github.com/microsoft/markitdown) - Microsoft 开源的文档转换库 