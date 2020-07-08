namespace ExcelPluginTest.ClosedXml
{
    using System.IO;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using ExcelPluginTest.ExportData;
    using ExcelPluginTest.Interfaces;

    public class ClosedXmlWordCreator : IWordCreator
    {
        private readonly ExportDataCreator _creator;

        public ClosedXmlWordCreator()
        {
            _creator = new ExportDataCreator();
        }

        public byte[] Create()
        {
            var data = _creator.Create(1);

            using (var stream = new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
                {
                    var mainDocument = wordDocument.AddMainDocumentPart();

                    mainDocument.Document = new Document();
                    mainDocument.Document.Body = new Body();

                    Paragraph p = new Paragraph();
                    // Run 1
                    Run r1 = new Run();
                    Text t1 = new Text("Pellentesque ") { Space = SpaceProcessingModeValues.Preserve };
                    // The Space attribute preserve white space before and after your text
                    r1.Append(t1);
                    p.Append(r1);

                    // Run 2 - Bold
                    Run r2 = new Run();
                    RunProperties rp2 = new RunProperties();
                    rp2.Bold = new Bold();
                    // Always add properties first
                    r2.Append(rp2);
                    Text t2 = new Text("commodo ") { Space = SpaceProcessingModeValues.Preserve };
                    r2.Append(t2);
                    p.Append(r2);

                    // Run 3
                    Run r3 = new Run();
                    Text t3 = new Text("rhoncus ") { Space = SpaceProcessingModeValues.Preserve };
                    r3.Append(t3);
                    p.Append(r3);

                    // Run 4 – Italic
                    Run r4 = new Run();
                    RunProperties rp4 = new RunProperties();
                    rp4.Italic = new Italic();
                    // Always add properties first
                    r4.Append(rp4);
                    Text t4 = new Text("mauris") { Space = SpaceProcessingModeValues.Preserve };
                    r4.Append(t4);
                    p.Append(r4);

                    // Run 5
                    Run r5 = new Run();
                    Text t5 = new Text(", sit ") { Space = SpaceProcessingModeValues.Preserve };
                    r5.Append(t5);
                    p.Append(r5);

                    // Run 6 – Italic , bold and underlined
                    Run r6 = new Run();
                    RunProperties rp6 = new RunProperties();
                    rp6.Italic = new Italic();
                    rp6.Bold = new Bold();
                    rp6.Underline = new Underline();
                    // Always add properties first
                    r6.Append(rp6);
                    Text t6 = new Text("amet ") { Space = SpaceProcessingModeValues.Preserve };
                    r6.Append(t6);
                    p.Append(r6);

                    // Run 7
                    Run r7 = new Run();
                    Text t7 = new Text("faucibus arcu ") { Space = SpaceProcessingModeValues.Preserve };
                    r7.Append(t7);
                    p.Append(r7);

                    // Run 8 – Red color
                    Run r8 = new Run();
                    RunProperties rp8 = new RunProperties();
                    rp8.Color = new Color() { Val = "FF0000" };
                    // Always add properties first
                    r8.Append(rp8);
                    Text t8 = new Text("porttitor ") { Space = SpaceProcessingModeValues.Preserve };
                    r8.Append(t8);
                    p.Append(r8);

                    // Run 9
                    Run r9 = new Run();
                    Text t9 = new Text("pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris.") { Space = SpaceProcessingModeValues.Preserve };
                    r9.Append(t9);
                    p.Append(r9);
                    // Add your paragraph to docx body

                    mainDocument.Document.Body.Append(p);
                }

                //var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
                //var mainDocument = document.AddMainDocumentPart();

                //AddStyles(mainDocument.AddNewPart<StyleDefinitionsPart>());

                //mainDocument.Document = new Document();
                //mainDocument.Document.Body = new Body();

                //AddParagraph(mainDocument.Document.Body, WordCreatorNames.Heading1Id, "This is a heading");

                //mainDocument.Document.Save();
                //document.Save();

                return stream.ToArray();
            }
        }

        private void AddStyles(StyleDefinitionsPart styleCollection)
        {
            var runProperties = new RunProperties();
            runProperties.Append(new Color { Val = "000000" });
            runProperties.Append(new RunFonts { Ascii = "Arial" });
            runProperties.Append(new Bold());
            runProperties.Append(new FontSize { Val = "28" });

            var style = new Style
                            {
                                StyleId = WordCreatorNames.Heading1Id,
                                Type = StyleValues.Paragraph,
                                CustomStyle = true
                            };
            style.Append(new StyleName { Val = WordCreatorNames.Heading1Id });
            style.Append(new BasedOn { Val = "Heading1" });
            style.Append(new NextParagraphStyle { Val = "Normal" });
            style.Append(runProperties);

            // we have to add style that we have created to the StylePart
            styleCollection.Styles = new Styles();
            styleCollection.Styles.Append(style);
            styleCollection.Styles.Save(); // we save the style part
        }

        private void AddParagraph(Body documentBody, string styleId, string value)
        {
            var paragraph = new Paragraph();
            paragraph.ParagraphProperties = new ParagraphProperties();
            paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleId };

            var run = new Run();
            var paragraphText = new Text(value);
            run.Append(paragraphText);
            paragraph.Append(run);

            documentBody.Append(paragraph);
        }
    }

    public static class WordCreatorNames
    {
        public const string Heading1Id = "TwentysixHeading1";

        public const string Heading1Name = "Twentysix Heading 1";
    }
}