using PicEx.Console.Utils;

namespace PicEx.Console.Extractors;

/// <summary>
/// Adobe Illustrator (.ai) 文件图片提取器
/// AI 文件本质是 PDF 格式，包含 PDF 兼容层
/// 直接复用 PDF 提取器即可
/// </summary>
public class AiImageExtractor
{
    /// <summary>
    /// 从 AI 文件中提取图片
    /// </summary>
    /// <param name="aiFilePath">AI 文件路径</param>
    /// <param name="outputDirectory">输出目录</param>
    public static void ExtractImagesFromAi(string aiFilePath, string outputDirectory)
    {
        try
        {
            System.Console.WriteLine("AI 文件检测到 PDF 兼容层，使用 PDF 提取器...");
            // AI 文件包含 PDF 兼容层，直接使用 PDF 提取器
            PdfImageExtractor.ExtractImagesFromPdf(aiFilePath, outputDirectory);
        }
        catch (Exception ex)
        {
            System.Console.WriteLine($"AI 文件解析失败（可能不包含 PDF 兼容层）: {ex.Message}");
            System.Console.WriteLine("提示：请在 Adobe Illustrator 中保存时启用 'Create PDF Compatible File' 选项");
            throw;
        }
    }
}
