# 构建阶段
FROM --platform=linux/amd64 mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# 复制项目文件
COPY PicEx.Console/*.csproj ./PicEx.Console/
COPY PicEx.Console/Extractors ./PicEx.Console/Extractors/
COPY PicEx.Console/Services ./PicEx.Console/Services/
COPY PicEx.Console/Utils ./PicEx.Console/Utils/
COPY PicEx.Console/Program.cs ./PicEx.Console/

# 还原依赖
RUN dotnet restore PicEx.Console/PicEx.Console.csproj

# 发布应用
RUN dotnet publish PicEx.Console/PicEx.Console.csproj -c Release -o out

# 运行阶段
FROM --platform=linux/amd64 mcr.microsoft.com/dotnet/runtime:8.0
WORKDIR /app

# 安装必要的系统依赖 (使用国内镜像源)
RUN sed -i 's/deb.debian.org/mirrors.aliyun.com/g' /etc/apt/sources.list.d/debian.sources && \
    apt-get update && apt-get install -y \
    libgdiplus \
    imagemagick \
    && rm -rf /var/lib/apt/lists/*

# 复制发布文件
COPY --from=build /app/out .

# 创建非 root 用户
RUN groupadd -r appuser && useradd -r -g appuser -u 1000 appuser

# 创建输入输出目录并设置权限
RUN mkdir -p /input /output && \
    chown -R appuser:appuser /input /output && \
    chmod -R 777 /input /output

# 切换到非 root 用户
USER appuser

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV DOTNET_CLI_TELEMETRY_OPTOUT=true

# 设置入口点和默认命令
ENTRYPOINT ["dotnet", "PicEx.Console.dll"]
CMD ["/input", "/output", "--watch"] 