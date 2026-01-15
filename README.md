# PicEx 图片提取工具

PicEx 是一个强大的文档图片提取工具，可以从各种文档格式中自动提取图片。支持轮询扫描模式，特别适合在 NAS 等环境中运行。

## ✨ 功能特点

- **支持多种文档格式：**
  - PDF 文件 (.pdf)
  - Word 文档 (.doc, .docx)
  - Excel 表格 (.xls, .xlsx)
  - PowerPoint 演示文稿 (.ppt, .pptx)
  - ZIP 压缩包 (.zip)
  - Adobe Illustrator 文件 (.ai) - 使用 PDF 兼容层
  - Adobe Photoshop 文件 (.psd) - 使用 ImageMagick 提取图层

- **智能文件处理：**
  - 🔄 **轮询模式** - 定时扫描目录（默认 30 秒），不依赖文件系统事件
  - 📁 支持子文件夹处理
  - 🔒 文件锁定检测，避免处理正在写入的文件
  - ⏭️ 智能去重，不重复处理未修改的文件
  - 📂 自动创建输出目录结构

- **容器化部署：**
  - 🐳 Docker 支持
  - 📦 Docker Compose 一键部署
  - 🚀 绿联云等 NAS 设备完美支持

## 📦 系统要求

- Docker 或 Docker Compose
- 至少 1GB 可用内存
- 至少 1GB 可用磁盘空间

## 🚀 快速开始

### 使用 Docker 镜像（推荐）

**阿里云镜像（国内推荐）：**
```bash
docker pull registry.cn-hangzhou.aliyuncs.com/zjsck/picex:latest
```

**使用 docker-compose：**
```bash
# 创建目录
mkdir -p input output

# 启动服务（轮询模式，每 30 秒扫描一次）
docker-compose up -d

# 查看日志
docker-compose logs -f
```

**直接运行：**
```bash
# 轮询模式（推荐）：每 30 秒扫描一次目录
docker run -d \
  --name picex \
  --user root \
  -v $(pwd)/input:/input \
  -v $(pwd)/output:/output \
  registry.cn-hangzhou.aliyuncs.com/zjsck/picex:latest /input /output 30

# 查看日志
docker logs -f picex
```

### 自定义轮询间隔

修改 `docker-compose.yml` 中的 `command` 参数：

```yaml
command: ["/input", "/output", "60"]  # 改为 60 秒
```

或在命令行中指定：

```bash
docker run ... picex /input /output 60  # 60 秒轮询间隔
```

### 从源码构建

```bash
# 克隆仓库
git clone https://github.com/zhangjunL/PicEx.git
cd PicEx

# 构建镜像
./build.sh

# 或使用 docker-compose
docker-compose build
```

## 📖 使用说明

### 工作模式

PicEx 使用**轮询扫描模式**：

1. 🔄 定时扫描输入目录（默认 30 秒）
2. 🔍 检测新文件或修改的文件
3. 📄 根据文件类型提取图片
4. 💾 保存到输出目录，保持原有目录结构

### 支持的格式

| 格式 | 扩展名 | 提取方式 |
|------|--------|----------|
| PDF | `.pdf` | 逐页提取嵌入图片 |
| Word | `.docx`, `.doc` | 提取嵌入媒体资源 |
| Excel | `.xlsx`, `.xls` | 提取嵌入图片 |
| PowerPoint | `.pptx`, `.ppt` | 提取嵌入媒体 |
| ZIP | `.zip` | 递归提取压缩包图片 |
| Adobe Illustrator | `.ai` | PDF 兼容层提取 |
| Adobe Photoshop | `.psd` | 图层提取（ImageMagick） |

### 输出结构

```
output/
├── document.pdf/
│   ├── image_page1_1.png
│   ├── image_page2_1.png
│   └── ...
├── design.psd/
│   ├── layer_1.png
│   ├── layer_2.png
│   └── ...
└── logo.ai/
    └── image_page1_1.png
```

## 🛠️ 技术栈

- **.NET 8.0** - 主开发框架
- **iText7** - PDF 处理
- **DocumentFormat.OpenXml** - Office 文档处理
- **ImageMagick** - PSD 图层提取
- **Docker** - 容器化部署

## 📁 项目结构

```
PicEx/
├── PicEx.Console/          # 主程序
│   ├── Extractors/        # 图片提取器
│   │   ├── PdfImageExtractor.cs
│   │   ├── ExcelImageExtractor.cs
│   │   ├── ArchiveImageExtractor.cs
│   │   ├── AiImageExtractor.cs
│   │   └── PsdImageExtractor.cs
│   ├── Services/          # 服务
│   │   └── FileWatcherService.cs
│   ├── Utils/             # 工具类
│   │   └── PathUtils.cs
│   └── Program.cs         # 入口程序
├── Dockerfile             # Docker 配置
├── docker-compose.yml     # Docker Compose 配置
├── build.sh               # 构建脚本
├── push-aliyun.sh         # 阿里云推送脚本
└── README.md
```

## 🔧 配置说明

### docker-compose.yml

```yaml
services:
  picex:
    image: registry.cn-hangzhou.aliyuncs.com/zjsck/picex:latest
    container_name: picex
    restart: unless-stopped
    user: root  # NAS 环境需要 root 权限
    volumes:
      - ./input:/input
      - ./output:/output
    command: ["/input", "/output", "30"]  # 30 秒轮询间隔
```

### 环境变量

- `DOTNET_CLI_TELEMETRY_OPTOUT=true` - 禁用 .NET 遥测

## 💡 使用场景

### 绿联云 NAS 部署

1. 通过绿联云 Docker 管理界面部署
2. 挂载 `input` 和 `output` 目录
3. 设置自动重启
4. 复制文件到 `input` 目录即可自动处理

### 批量处理文档

```bash
# 复制所有 PDF 到 input 目录
cp ~/Documents/*.pdf ./input/

# 等待最多 30 秒（轮询间隔）
sleep 30

# 查看输出
ls -R ./output/
```

### 定时任务

配合 cron 或 NAS 定时任务，定期运行容器处理文件：

```bash
# 每天凌晨 2 点运行
0 2 * * * cd /path/to/PicEx && docker-compose up -d && sleep 3600 && docker-compose down
```

## ⚠️ 注意事项

1. **文件权限**：确保 Docker 容器有读写权限
   - NAS 环境建议使用 `user: root`
   - 或设置 output 目录权限为 777

2. **轮询间隔**：根据需求调整
   - 10-30 秒：快速响应，稍占资源
   - 60-300 秒：低资源占用

3. **PSD 文件**：
   - 复杂的 PSD 可能处理时间较长
   - 大型 PSD 建议增加容器内存限制

4. **AI 文件**：
   - 需要包含 PDF 兼容层
   - 在 Illustrator 保存时启用 "Create PDF Compatible File"

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT License

