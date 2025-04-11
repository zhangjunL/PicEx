namespace PicEx.Console.Utils;

public static class PathUtils
{
    public static string GetSafeFileName(string filePath)
    {
        string fileName = Path.GetFileNameWithoutExtension(filePath);
        // 替换非法字符
        char[] invalidChars = Path.GetInvalidFileNameChars();
        foreach (char c in invalidChars)
        {
            fileName = fileName.Replace(c, '_');
        }
        return fileName;
    }

    public static string CreateOutputDirectory(string baseDir, string filePath)
    {
        string safeFileName = GetSafeFileName(filePath);
        string outputDir = Path.Combine(baseDir, safeFileName);
        Directory.CreateDirectory(outputDir);
        return outputDir;
    }
} 