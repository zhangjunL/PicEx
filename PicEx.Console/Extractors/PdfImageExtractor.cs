using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using PicEx.Console.Utils;

namespace PicEx.Console.Extractors;

public class PdfImageExtractor
{
    public static void ExtractImagesFromPdf(string pdfPath, string outputDirectory)
    {
        // 创建以PDF文件名命名的子文件夹
        string outputDir = PathUtils.CreateOutputDirectory(outputDirectory, pdfPath);
        
        PdfReader reader = new PdfReader(pdfPath);
        PdfDocument document = new PdfDocument(reader);

        for (int pageNumber = 1; pageNumber <= document.GetNumberOfPages(); pageNumber++)
        {
            PdfPage page = document.GetPage(pageNumber);
            ImageExtractionListener listener = new ImageExtractionListener(outputDir, pageNumber);
            PdfCanvasProcessor processor = new PdfCanvasProcessor(listener);
            processor.ProcessPageContent(page);

            processor.Reset();
        }

        document.Close();
        System.Console.WriteLine($"PDF图片已提取到: {outputDir}");
    }
}

public class ImageExtractionListener : IEventListener
{
    private string outputDirectory;
    private int pageNumber;
    private int imageCounter;

    public ImageExtractionListener(string outputDirectory, int pageNumber)
    {
        this.outputDirectory = outputDirectory;
        this.pageNumber = pageNumber;
        this.imageCounter = 1;
    }

    public void EventOccurred(IEventData data, EventType type)
    {
        if (type == EventType.RENDER_IMAGE)
        {
            ImageRenderInfo imageRenderInfo = (ImageRenderInfo)data;
            PdfImageXObject imageObject = imageRenderInfo.GetImage();
            string imageName = $"{outputDirectory}/image_page{pageNumber}_{imageCounter}.png";
            using (FileStream stream = new FileStream(imageName, FileMode.Create))
            {
                stream.Write(imageObject.GetImageBytes());
            }
            imageCounter++;
        }
    }

    public ICollection<EventType> GetSupportedEvents()
    {
        return new List<EventType> { EventType.RENDER_IMAGE };
    }
} 