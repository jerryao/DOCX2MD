FROM python:3.10-slim

WORKDIR /app

# 安装依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用程序代码
COPY . .

# 确保上传和下载文件夹存在
RUN mkdir -p uploads downloads

# 暴露端口
EXPOSE 5000

# 设置环境变量
ENV FLASK_APP=app.py
ENV PYTHONUNBUFFERED=1
ENV SECRET_KEY=change-me-in-production

# 设置用户
RUN useradd -m appuser
RUN chown -R appuser:appuser /app
USER appuser

# 启动应用
CMD ["python", "app.py"] 