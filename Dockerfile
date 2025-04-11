FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# 复制项目文件
COPY PicEx.Console/*.csproj ./PicEx.Console/
COPY PicEx.Console/Extractors ./PicEx.Console/Extractors/
COPY PicEx.Console/Services ./PicEx.Console/Services/
COPY PicEx.Console/Utils ./PicEx.Console/Utils/
COPY PicEx.Console/Program.cs ./PicEx.Console/

# 还原包
RUN dotnet restore PicEx.Console/PicEx.Console.csproj

# 发布应用
RUN dotnet publish PicEx.Console/PicEx.Console.csproj -c Release -o out

# 构建运行时镜像
FROM mcr.microsoft.com/dotnet/runtime:8.0
WORKDIR /app

# 安装必要的系统依赖
RUN apt-get update && apt-get install -y \
    libgdiplus \
    && rm -rf /var/lib/apt/lists/*

# 从构建阶段复制发布文件
COPY --from=build /app/out .

# 设置入口点
ENTRYPOINT ["dotnet", "PicEx.Console.dll"] 