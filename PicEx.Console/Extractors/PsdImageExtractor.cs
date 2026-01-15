using ImageMagick;
using PicEx.Console.Utils;

namespace PicEx.Console.Extractors;

/// <summary>
/// Adobe Photoshop (.psd) 文件图片提取器
/// 使用 ImageMagick 库读取 PSD 文件并提取图层
/// </summary>
public class PsdImageExtractor
{
    /// <summary>
    /// 从 PSD 文件中提取图片
    /// </summary>
    /// <param name="psdPath">PSD 文件路径</param>
    /// <param name="outputDirectory">输出目录</param>
    public static void ExtractImagesFromPsd(string psdPath, string outputDirectory)
    {
        string outputDir = PathUtils.CreateOutputDirectory(outputDirectory, psdPath);

        try
        {
            System.Console.WriteLine($"正在读取 PSD 文件: {Path.GetFileName(psdPath)}");

            // 使用 ImageMagick 读取 PSD
            using var images = new MagickImageCollection();
            images.Read(psdPath);

            System.Console.WriteLine($"  检测到 {images.Count} 个图层/图像");

            int imageIndex = 1;

            foreach (var image in images)
            {
                try
                {
                    // 生成文件名
                    string outputPath = Path.Combine(outputDir, $"layer_{imageIndex}.png");

                    // 保存为 PNG
                    image.Write(outputPath);

                    System.Console.WriteLine($"  已提取: layer_{imageIndex}.png ({image.Width}x{image.Height})");
                    imageIndex++;
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine($"  警告: 无法保存图层 {imageIndex}: {ex.Message}");
                }
            }

            System.Console.WriteLine($"PSD 图片已提取到: {outputDir}");
            System.Console.WriteLine($"成功提取 {imageIndex - 1} 个图层");
        }
        catch (Exception ex)
        {
            System.Console.WriteLine($"处理 PSD 文件时出错: {ex.Message}");
            throw;
        }
    }
}
