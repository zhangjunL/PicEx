using PicEx.Console.Extractors;
using PicEx.Console.Services;
using PicEx.Console.Utils;

namespace PicEx.Console;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            System.Console.WriteLine("用法: PicEx.Console <输入目录> <输出目录>");
            return;
        }

        string inputDirectory = args[0];
        string outputDirectory = args[1];

        if (!Directory.Exists(inputDirectory))
        {
            System.Console.WriteLine($"输入目录不存在: {inputDirectory}");
            return;
        }

        Directory.CreateDirectory(outputDirectory);

        var watcher = new FileWatcherService(inputDirectory, outputDirectory);
        watcher.Start();

        System.Console.CancelKeyPress += (sender, e) =>
        {
            e.Cancel = true;
            watcher.Stop();
            System.Console.WriteLine("正在停止监视...");
            Environment.Exit(0);
        };

        // 保持程序运行
        while (true)
        {
            Thread.Sleep(1000);
        }
    }

    static void ProcessSingleFile(string filePath, string outputDir)
    {
        string fileType = Path.GetExtension(filePath).ToLower();

        switch (fileType)
        {
            case ".pdf":
                PdfImageExtractor.ExtractImagesFromPdf(filePath, outputDir);
                break;
            case ".xlsx":
                ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDir, "xl");
                break;
            case ".docx":
                ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDir, "word");
                break;
            case ".pptx":
                ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDir, "ppt");
                break;
            case ".zip":
                ArchiveImageExtractor.ExtractImagesFromArchive(filePath, outputDir);
                break;
            default:
                System.Console.WriteLine($"不支持的文件类型: {fileType}");
                ShowHelp();
                break;
        }
    }

    static void ShowHelp()
    {
        System.Console.WriteLine("使用方法:");
        System.Console.WriteLine("  单文件处理: PicEx.Console <文件路径> <输出目录>");
        System.Console.WriteLine("  监控模式: PicEx.Console <输入目录> <输出目录> --watch");
        System.Console.WriteLine("\n支持的文件类型:");
        System.Console.WriteLine("  - PDF文件 (.pdf)");
        System.Console.WriteLine("  - Excel文件 (.xlsx)");
        System.Console.WriteLine("  - Word文件 (.docx)");
        System.Console.WriteLine("  - PowerPoint文件 (.pptx)");
        System.Console.WriteLine("  - 压缩文件 (.zip)");
    }
}
