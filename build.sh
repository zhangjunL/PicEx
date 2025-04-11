#!/bin/bash

# 构建Docker镜像
docker build -t picex .

# 显示使用说明
echo "构建完成！"
echo "使用方法:"
echo "1. 单文件处理:"
echo "   docker run -v <输入文件目录>:/input -v <输出目录>:/output picex /input/文件名 /output"
echo ""
echo "2. 监控模式:"
echo "   docker run -v <输入目录>:/input -v <输出目录>:/output picex /input /output --watch"
echo ""
echo "示例:"
echo "1. 处理单个文件:"
echo "   docker run -v ./input:/input -v ./output:/output picex /input/document.pdf /output"
echo ""
echo "2. 监控目录:"
echo "   docker run -v ./input:/input -v ./output:/output picex /input /output --watch" 