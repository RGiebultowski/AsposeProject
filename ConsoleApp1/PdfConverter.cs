using Aspose.Words;

namespace ConsoleApp1
{
    public class PdfConverter
    {
        public void ConvertToPdf(string docPath, string pdfPath)
        {
            Document doc = new Document(docPath);
            doc.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}
