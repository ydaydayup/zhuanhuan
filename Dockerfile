# 使用Python 3.9作为基础镜像
FROM python:3.12-slim-bookworm

# 设置工作目录
WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends curl ca-certificates
COPY . .

RUN pip install uv
# 安装Python依赖
RUN uv sync --frozen --no-cache

# 创建必要的目录
RUN mkdir -p uploads results metadata

# 设置环境变量
ENV PYTHONUNBUFFERED=1

# 暴露端口
EXPOSE 5000

# 启动命令
#CMD ["python", "app.py", "--host", "0.0.0.0"]