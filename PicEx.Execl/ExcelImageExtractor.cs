using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;

namespace PicEx.Execl
{
    public class ExcelImageExtractor
    {
        public static void ExtractImagesFromExcel(string excelFilePath, string outputFolderPath)
        {
            string tempFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());

            // 创建临时文件夹
            Directory.CreateDirectory(tempFolderPath);

            // 将原文件复制到临时文件夹
            string tempExcelFilePath = Path.Combine(tempFolderPath, Path.GetFileName(excelFilePath));
            File.Copy(excelFilePath, tempExcelFilePath);

            // 修改临时文件后缀名为 RAR
            string tempRarFilePath = Path.ChangeExtension(tempExcelFilePath, ".rar");
            File.Move(tempExcelFilePath, tempRarFilePath);

            // 将 RAR 文件解压到临时文件夹
            ZipFile.ExtractToDirectory(tempRarFilePath, tempFolderPath);

            // 获取包含图片的文件夹路径
            string imagesFolderPath = Path.Combine(tempFolderPath, "xl","media");

            if (Directory.Exists(imagesFolderPath))
            {
                // 创建保存图片的目标文件夹
                Directory.CreateDirectory(outputFolderPath);

                // 遍历图片文件夹中的所有文件
                foreach (string imagePath in Directory.GetFiles(imagesFolderPath))
                {
                    // 读取图片数据
                    byte[] imageData = File.ReadAllBytes(imagePath);

                    // 构造目标文件路径
                    string outputFilePath = Path.Combine(outputFolderPath, Path.GetFileName(imagePath));

                    // 如果目标文件已经存在，则在文件名后追加数字序号
                    if (File.Exists(outputFilePath))
                    {
                        string fileName = Path.GetFileNameWithoutExtension(outputFilePath);
                        string fileExtension = Path.GetExtension(outputFilePath);
                        int index = 1;

                        do
                        {
                            string indexedFileName = $"{fileName}_{index}{fileExtension}";
                            outputFilePath = Path.Combine(outputFolderPath, indexedFileName);
                            index++;
                        }
                        while (File.Exists(outputFilePath));
                    }

                    // 保存图片到文件
                    File.WriteAllBytes(outputFilePath, imageData);

                    Console.WriteLine("图片提取成功，保存路径为：" + outputFilePath);
                }
            }
            else
            {
                Console.WriteLine("未找到包含图片的文件夹。");
            }

            // 清除临时文件和文件夹
            File.Delete(tempRarFilePath);
            Directory.Delete(tempFolderPath, true);
        }
    }
}