using System.IO;
using PicEx.Console.Extractors;
using PicEx.Console.Utils;

namespace PicEx.Console.Services;

public class FileWatcherService
{
    private readonly FileSystemWatcher _watcher;
    private readonly string _outputDirectory;
    private readonly Dictionary<string, DateTime> _processedFiles;

    public FileWatcherService(string inputDirectory, string outputDirectory)
    {
        _outputDirectory = outputDirectory;
        _processedFiles = new Dictionary<string, DateTime>();

        _watcher = new FileSystemWatcher(inputDirectory)
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.DirectoryName,
            Filter = "*.*",
            EnableRaisingEvents = false,
            IncludeSubdirectories = true
        };

        _watcher.Created += OnFileCreated;
        _watcher.Changed += OnFileChanged;
    }

    public void Start()
    {
        _watcher.EnableRaisingEvents = true;
        System.Console.WriteLine($"开始监视目录: {_watcher.Path}");
        System.Console.WriteLine("支持的文件类型: PDF, Word, Excel, PowerPoint, ZIP");
        System.Console.WriteLine("按 Ctrl+C 停止监视");
    }

    public void Stop()
    {
        _watcher.EnableRaisingEvents = false;
        _watcher.Dispose();
    }

    private void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        ProcessFile(e.FullPath);
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        ProcessFile(e.FullPath);
    }

    private void ProcessFile(string filePath)
    {
        try
        {
            if (IsFileLocked(filePath))
            {
                return;
            }

            var fileInfo = new FileInfo(filePath);
            if (!IsSupportedFileType(fileInfo.Extension))
            {
                return;
            }

            if (_processedFiles.TryGetValue(filePath, out var lastProcessed))
            {
                if ((DateTime.Now - lastProcessed).TotalSeconds < 5)
                {
                    return;
                }
            }

            var relativePath = Path.GetRelativePath(_watcher.Path, fileInfo.DirectoryName!);
            var outputDir = Path.Combine(_outputDirectory, relativePath);
            Directory.CreateDirectory(outputDir);

            var outputPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(fileInfo.Name));
            System.Console.WriteLine($"正在处理: {filePath}");

            switch (fileInfo.Extension.ToLower())
            {
                case ".pdf":
                    PdfImageExtractor.ExtractImagesFromPdf(filePath, outputPath);
                    break;
                case ".docx":
                case ".doc":
                    ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputPath, "word");
                    break;
                case ".xlsx":
                case ".xls":
                    ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputPath, "xl");
                    break;
                case ".pptx":
                case ".ppt":
                    ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputPath, "ppt");
                    break;
                case ".zip":
                    ArchiveImageExtractor.ExtractImagesFromArchive(filePath, outputPath);
                    break;
            }

            _processedFiles[filePath] = DateTime.Now;
            System.Console.WriteLine($"处理完成: {filePath}");
        }
        catch (Exception ex)
        {
            System.Console.WriteLine($"处理文件时出错 {filePath}: {ex.Message}");
        }
    }

    private bool IsSupportedFileType(string extension)
    {
        return extension.ToLower() switch
        {
            ".pdf" or ".docx" or ".doc" or ".xlsx" or ".xls" or ".pptx" or ".ppt" or ".zip" => true,
            _ => false
        };
    }

    private bool IsFileLocked(string filePath)
    {
        try
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
    }
} 