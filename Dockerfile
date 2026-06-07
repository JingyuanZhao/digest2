FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PORT=7860
ENV FLASK_ENV=production
ENV PYTHONUNBUFFERED=1

# 安装编译工具和依赖
RUN apt-get update && apt-get install -y \
    gcc \
    make \
    libc-dev \
    && rm -rf /var/lib/apt/lists/*

# 设置 C 编译器标志以兼容 digest2 的旧代码
ENV CFLAGS="-Wno-implicit-function-declaration -Wno-error=implicit-function-declaration"

# 复制依赖文件
COPY requirements.txt .

# 安装依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY . .

# 暴露端口
EXPOSE 7860

# 启动命令
CMD ["gunicorn", "--workers", "2", "--bind", "0.0.0.0:7860", "--timeout", "120", "app:application"]
