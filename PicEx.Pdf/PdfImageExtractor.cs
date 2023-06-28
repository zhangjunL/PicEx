using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
namespace PicEx.Pdf;

public class PdfImageExtractor
{
    public static void ExtractImagesFromPdf(string pdfPath, string outputDirectory)
    {
        PdfReader reader = new PdfReader(pdfPath);
        PdfDocument document = new PdfDocument(reader);

        for (int pageNumber = 1; pageNumber <= document.GetNumberOfPages(); pageNumber++)
        {
            PdfPage page = document.GetPage(pageNumber);
            ImageExtractionListener listener = new ImageExtractionListener(outputDirectory, pageNumber);
            PdfCanvasProcessor processor = new PdfCanvasProcessor(listener);
            processor.ProcessPageContent(page);

            processor.Reset();
        }

        document.Close();
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

