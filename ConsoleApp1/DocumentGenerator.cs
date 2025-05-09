using Aspose.Words;

namespace ConsoleApp1
{
    public class DocumentGenerator
    {
        private readonly Document doc;
        private readonly DocumentBuilder builder;

        public DocumentGenerator()
        {
            doc = new Document();
            builder = new DocumentBuilder(doc);
        }

        public string Generate(string outputPath)
        {
            AddTitle("Test Doc.");

            for (int i = 1; i <= 3; i++)
            {
                AddSection($"Section {i}");
                AddLoremParagraph();
            }

            doc.Save(outputPath);
            return outputPath;
        }

        private void AddTitle(string title)
        {
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 16;
            builder.Font.Bold = true;
            builder.Writeln(title);
            builder.InsertParagraph();
        }

        private void AddSection(string heading)
        {
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Font.Size = 14;
            builder.Font.Bold = true;
            builder.Writeln(heading);
        }

        private void AddLoremParagraph()
        {
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            builder.Font.Bold = false;
            builder.Font.Size = 12;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                             "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.InsertParagraph();
        }
    }
}
