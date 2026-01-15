using PicEx.Console.Extractors;
using PicEx.Console.Utils;

namespace PicEx.Console;

class Program
{
    private static string _inputDirectory = "";
    private static string _outputDirectory = "";
    private static Dictionary<string, DateTime> _processedFiles = new();
    private static bool _isRunning = true;

    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            System.Console.WriteLine("用法: PicEx.Console <输入目录> <输出目录> [轮询间隔秒数]");
            System.Console.WriteLine("示例: PicEx.Console /input /output 30");
            return;
        }

        _inputDirectory = args[0];
        _outputDirectory = args[1];

        // 轮询间隔（秒），默认 30 秒
        int pollInterval = args.Length > 2 && int.TryParse(args[2], out int interval) ? interval : 30;

        if (!Directory.Exists(_inputDirectory))
        {
            System.Console.WriteLine($"输入目录不存在: {_inputDirectory}");
            return;
        }

        Directory.CreateDirectory(_outputDirectory);

        System.Console.WriteLine($"=========================================");
        System.Console.WriteLine($"PicEx 图片提取工具 - 轮询模式");
        System.Console.WriteLine($"=========================================");
        System.Console.WriteLine($"输入目录: {_inputDirectory}");
        System.Console.WriteLine($"输出目录: {_outputDirectory}");
        System.Console.WriteLine($"轮询间隔: {pollInterval} 秒");
        System.Console.WriteLine($"支持的文件类型: PDF, Word, Excel, PowerPoint, ZIP");
        System.Console.WriteLine($"按 Ctrl+C 停止");
        System.Console.WriteLine($"=========================================");

        // 启动时先扫描一次
        System.Console.WriteLine($"开始首次扫描...");
        ScanAndProcessFiles();

        // 设置取消按键
        System.Console.CancelKeyPress += (sender, e) =>
        {
            e.Cancel = true;
            _isRunning = false;
            System.Console.WriteLine("\n正在停止...");
            Environment.Exit(0);
        };

        // 定时轮询
        while (_isRunning)
        {
            Thread.Sleep(pollInterval * 1000);
            if (_isRunning)
            {
                ScanAndProcessFiles();
            }
        }
    }

    static void ScanAndProcessFiles()
    {
        try
        {
            var files = Directory.GetFiles(_inputDirectory, "*.*", SearchOption.AllDirectories);
            int processedCount = 0;

            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);

                if (!IsSupportedFileType(fileInfo.Extension))
                {
                    continue;
                }

                // 检查文件是否已处理且未修改
                if (_processedFiles.TryGetValue(file, out DateTime lastProcessed))
                {
                    if (fileInfo.LastWriteTime <= lastProcessed)
                    {
                        continue; // 文件未修改，跳过
                    }
                }

                // 处理文件
                ProcessFile(file);
                _processedFiles[file] = fileInfo.LastWriteTime;
                processedCount++;
            }

            if (processedCount > 0)
            {
                System.Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 已处理 {processedCount} 个文件");
            }
        }
        catch (Exception ex)
        {
            System.Console.WriteLine($"扫描目录时出错: {ex.Message}");
        }
    }

    static void ProcessFile(string filePath)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            var relativePath = Path.GetRelativePath(_inputDirectory, fileInfo.DirectoryName!);
            var outputDir = Path.Combine(_outputDirectory, relativePath);
            Directory.CreateDirectory(outputDir);

            var outputPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(fileInfo.Name));
            System.Console.WriteLine($"正在处理: {Path.GetFileName(filePath)}");

            switch (fileInfo.Extension.ToLower())
            {
                case ".pdf":
                case ".ai":  // AI 文件本质是 PDF，使用 PDF 提取器
                    PdfImageExtractor.ExtractImagesFromPdf(filePath, outputDir);
                    break;
                case ".psd":  // Adobe Photoshop 文件
                    PsdImageExtractor.ExtractImagesFromPsd(filePath, outputDir);
                    break;
                case ".docx":
                case ".doc":
                    ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDir, "word");
                    break;
                case ".xlsx":
                case ".xls":
                    ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDir, "xl");
                    break;
                case ".pptx":
                case ".ppt":
                    ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDir, "ppt");
                    break;
                case ".zip":
                    ArchiveImageExtractor.ExtractImagesFromArchive(filePath, outputDir);
                    break;
            }

            System.Console.WriteLine($"处理完成: {Path.GetFileName(filePath)}");
        }
        catch (Exception ex)
        {
            System.Console.WriteLine($"处理文件时出错 {Path.GetFileName(filePath)}: {ex.Message}");
        }
    }

    static bool IsSupportedFileType(string extension)
    {
        return extension.ToLower() switch
        {
            ".pdf" or ".docx" or ".doc" or ".xlsx" or ".xls" or ".pptx" or ".ppt" or ".zip"
            or ".ai" or ".psd" => true,  // 新增 AI 和 PSD 支持
            _ => false
        };
    }
}
