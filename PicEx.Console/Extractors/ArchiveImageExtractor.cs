using System.IO.Compression;
using PicEx.Console.Utils;

namespace PicEx.Console.Extractors;

public class ArchiveImageExtractor
{
    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };

    public static void ExtractImagesFromArchive(string archivePath, string outputDirectory)
    {
        // 创建以压缩文件名命名的子文件夹
        string outputDir = PathUtils.CreateOutputDirectory(outputDirectory, archivePath);
        
        string tempFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempFolderPath);

        try
        {
            // 解压文件
            ZipFile.ExtractToDirectory(archivePath, tempFolderPath);

            // 递归搜索图片文件
            SearchAndExtractImages(tempFolderPath, outputDir);

            System.Console.WriteLine($"成功从压缩文件中提取图片到: {outputDir}");
        }
        catch (Exception ex)
        {
            System.Console.WriteLine($"处理压缩文件时出错: {ex.Message}");
        }
        finally
        {
            // 清理临时文件
            if (Directory.Exists(tempFolderPath))
            {
                Directory.Delete(tempFolderPath, true);
            }
        }
    }

    private static void SearchAndExtractImages(string sourceDir, string targetDir)
    {
        // 处理当前目录中的文件
        foreach (string file in Directory.GetFiles(sourceDir))
        {
            string extension = Path.GetExtension(file).ToLower();
            if (ImageExtensions.Contains(extension))
            {
                string fileName = Path.GetFileName(file);
                string targetPath = Path.Combine(targetDir, fileName);

                // 如果文件已存在，添加序号
                if (File.Exists(targetPath))
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    string ext = Path.GetExtension(fileName);
                    int index = 1;

                    do
                    {
                        targetPath = Path.Combine(targetDir, $"{nameWithoutExt}_{index}{ext}");
                        index++;
                    } while (File.Exists(targetPath));
                }

                File.Copy(file, targetPath);
                System.Console.WriteLine($"提取图片: {fileName}");
            }
        }

        // 递归处理子目录
        foreach (string dir in Directory.GetDirectories(sourceDir))
        {
            SearchAndExtractImages(dir, targetDir);
        }
    }
} 