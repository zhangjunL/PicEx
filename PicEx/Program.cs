// See https://aka.ms/new-console-template for more information
using PicEx.Execl;
using PicEx.Pdf;
using PicEx.Word;


string pdfPath = @"C:\Users\Administrator\Desktop\pdftest\3.xlsx";
string outputDirectory = @"C:\Users\Administrator\Desktop\pdftest\3";
// 从pdfPath中获取文件扩展名
string extension = Path.GetExtension(pdfPath);
// 如果扩展名不是pdf，则抛出异常
switch (extension)
{
    case ".pdf":
        PdfImageExtractor.ExtractImagesFromPdf(pdfPath, outputDirectory);
        break;
    case ".docx":
        WordImageExtractor.ExtractImagesFromWord(pdfPath, outputDirectory);
        break;
    case ".xlsx":
        ExcelImageExtractor.ExtractImagesFromExcel(pdfPath, outputDirectory);
        break;
    default:
        Console.WriteLine("该文件类型暂不支持解析");
        break;
        //throw new ArgumentException("该文件类型暂不支持解析");
        
}


Console.WriteLine("Press any key to exit");
Console.ReadKey();

