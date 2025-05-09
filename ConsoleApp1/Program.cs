using ConsoleApp1;

internal class Program
{
    private static void Main(string[] args)
    {
        string outputDir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Output");
        string formattedDate = DateTime.Now.ToString("yyyyMMdd");
        Directory.CreateDirectory(outputDir);

        string docPath = Path.Combine(outputDir,$"_{formattedDate}.docx");
        string pdfPath = Path.Combine(outputDir, $"_{formattedDate}.pdf");

        Console.WriteLine("Generating doc...");

        try
        {
            var generator = new DocumentGenerator();
            string generatedDocPath = generator.Generate(docPath);
            Console.WriteLine($"Word file generated: {generatedDocPath}");

            var converter = new PdfConverter();
            converter.ConvertToPdf(generatedDocPath, pdfPath);
            Console.WriteLine($"PDF file created: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Exception:");
            Console.Error.WriteLine(ex.Message);
        }

        Console.WriteLine("Success... Hit any button to exit.");
        Console.ReadKey();
    }
}