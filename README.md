# PicEx 图片导出工具

PicEx 是一个强大的文档图片提取工具,可以从各种文档格式中提取图片。它支持监控模式,可以自动处理指定目录中的文件。

## 功能特点

- 支持多种文档格式:
  - PDF 文件 (.pdf)
  - Word 文档 (.doc, .docx)
  - Excel 表格 (.xls, .xlsx)
  - PowerPoint 演示文稿 (.ppt, .pptx)
  - ZIP 压缩包 (.zip)
- 支持子文件夹处理
- 支持文件监控模式
- 支持单文件处理模式
- 自动创建输出目录结构
- 防止重复处理同一文件
- 文件锁定检测

## 系统要求

- Docker
- 至少 1GB 可用内存
- 至少 1GB 可用磁盘空间

## 安装说明

### 1. 使用 Docker 镜像

从 GitHub Container Registry 拉取镜像:

```bash
docker pull ghcr.io/yourusername/picex:latest
```

### 2. 从源码构建

1. 克隆仓库:
```bash
git clone https://github.com/yourusername/PicEx.git
cd PicEx
```

2. 构建 Docker 镜像:
```bash
./build.sh
```

## 使用方法

### 1. 监控模式

监控指定目录,自动处理新文件:

```bash
docker run -v <输入目录>:/input -v <输出目录>:/output picex /input /output
```

示例:
```bash
# 监控 input 目录,输出到 output 目录
docker run -v ./input:/input -v ./output:/output picex /input /output
```

### 2. 单文件处理模式

处理单个文件:

```bash
docker run -v <输入文件目录>:/input -v <输出目录>:/output picex /input/文件名 /output
```

示例:
```bash
# 处理单个 PDF 文件
docker run -v ./input:/input -v ./output:/output picex /input/document.pdf /output
```

### 3. 使用 Docker Compose

也可以使用 Docker Compose 来运行:

```bash
# 启动服务
docker-compose up

# 在后台运行
docker-compose up -d
```

## 开发环境

### 1. 使用 Docker Compose 开发

项目提供了开发环境的 Docker Compose 配置:

```bash
# 启动开发环境
docker-compose up

# 在容器中运行命令
docker-compose exec picex-dev dotnet build
docker-compose exec picex-dev dotnet test
```

开发环境配置:
- 源代码目录挂载到容器的 /app 目录
- 禁用 .NET CLI 遥测
- 使用 .NET 8.0 SDK 镜像

### 2. 本地开发

也可以直接在本地开发:

1. 安装 .NET 8.0 SDK
2. 克隆仓库
3. 使用 Visual Studio 或 VS Code 打开解决方案
4. 运行 `dotnet restore` 还原包
5. 运行 `dotnet build` 构建项目
6. 运行 `dotnet test` 运行测试

## CI/CD

项目使用 GitHub Actions 进行持续集成和部署:

1. 自动构建 Docker 镜像
2. 推送到 GitHub Container Registry
3. 支持以下触发条件:
   - 推送到 main/master 分支
   - 创建版本标签 (v*.*.*)
   - 创建 Pull Request
   - 手动触发

镜像标签规则:
- 分支名称 (例如: main, master)
- 版本号 (例如: v1.0.0)
- 主版本号.次版本号 (例如: v1.0)
- Git commit SHA (短格式)

## 目录结构

```
PicEx/
├── input/              # 输入目录
├── output/            # 输出目录
├── PicEx.Console/     # 主程序
│   ├── Extractors/   # 图片提取器
│   ├── Services/     # 服务
│   └── Utils/        # 工具类
├── PicEx.win/        # Windows 平台特定代码
├── PicEx.Excel/      # Excel 处理相关代码
├── PicEx.Pdf/        # PDF 处理相关代码
├── Dockerfile         # Docker 配置
├── docker-compose.yml # Docker Compose 配置
├── build.sh          # 构建脚本
└── README.md         # 说明文档
```

## 输出说明

- 提取的图片将保存在输出目录中
- 输出目录会保持与输入目录相同的结构
- 图片文件名格式: `原文件名_序号.扩展名`
- 支持的图片格式: PNG, JPEG, GIF, BMP

## 注意事项

1. 确保输入目录和输出目录有足够的读写权限
2. 大文件处理可能需要较长时间
3. 建议定期清理输出目录
4. 程序会自动跳过不支持的文件类型
5. Windows 平台可能需要额外的配置

## 开发说明

- 使用 .NET 8.0 开发
- 主要依赖:
  - iText7 (PDF 处理)
  - DocumentFormat.OpenXml (Office 文档处理)
  - System.IO.Compression (ZIP 处理)
- 项目结构:
  - PicEx.Console: 主程序
  - PicEx.win: Windows 平台特定功能
  - PicEx.Excel: Excel 文件处理
  - PicEx.Pdf: PDF 文件处理

## 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目。

## 许可证

MIT License

