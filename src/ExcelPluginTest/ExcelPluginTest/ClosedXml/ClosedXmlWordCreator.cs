namespace ExcelPluginTest.ClosedXml
{
    using System;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using ExcelPluginTest.ExportData;
    using ExcelPluginTest.Extensions;
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

                    AddTable(mainDocument.Document.Body, data.First());
                }

                return stream.ToArray();
            }
        }

        private void AddTable(Body docBody, ExportDataRecord data)
        {
            var table = new Table();
            var tableProperties = new TableProperties();
            var tableStyle = new TableStyle { Val = "TableGrid" };
            var tableWidth = new TableWidth { Width = "1500", Type = TableWidthUnitValues.Pct};
            var tableBorders = new TableBorders
                                   {
                                       TopBorder = new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines) },
                                       BottomBorder = new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines) },
                                       LeftBorder = new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines) },
                                       RightBorder = new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines) },
                                       InsideHorizontalBorder = new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines) },
                                       InsideVerticalBorder = new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines) }
                                   };
            tableProperties.Append(tableStyle, tableWidth, tableBorders);
            table.Append(tableProperties);

            AddTableRow(table, nameof(ExportDataRecord.Title), data.Title);
            AddTableRow(table, nameof(ExportDataRecord.Forename), data.Forename);
            AddTableRow(table, nameof(ExportDataRecord.Surname), data.Surname);
            AddTableRow(table, nameof(ExportDataRecord.PreviousSurname), data.PreviousSurname);
            AddTableRow(table, nameof(ExportDataRecord.Profession), data.Profession);
            AddTableRow(table, nameof(ExportDataRecord.Number), data.Number);
            AddTableRow(table, nameof(ExportDataRecord.NumberReference), data.NumberReference);
            AddTableRow(table, nameof(ExportDataRecord.NumberDate), data.NumberDate);
            AddTableRow(table, nameof(ExportDataRecord.PreviouslyProvided), data.PreviouslyProvided);
            AddTableRow(table, nameof(ExportDataRecord.PreviouslyProvidedReference), data.PreviouslyProvidedReference);
            AddTableRow(table, nameof(ExportDataRecord.CategoryOne), data.CategoryOne);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwo), data.CategoryTwo);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThree), data.CategoryThree);
            AddTableRow(table, nameof(ExportDataRecord.CategoryFour), data.CategoryFour);
            AddTableRow(table, nameof(ExportDataRecord.CategoryFive), data.CategoryFive);
            AddTableRow(table, nameof(ExportDataRecord.CategorySix), data.CategorySix);
            AddTableRow(table, nameof(ExportDataRecord.CategorySeven), data.CategorySeven);
            AddTableRow(table, nameof(ExportDataRecord.CategoryEight), data.CategoryEight);
            AddTableRow(table, nameof(ExportDataRecord.CategoryNine), data.CategoryNine);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTen), data.CategoryTen);
            AddTableRow(table, nameof(ExportDataRecord.CategoryEleven), data.CategoryEleven);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwelve), data.CategoryTwelve);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirteen), data.CategoryThirteen);
            AddTableRow(table, nameof(ExportDataRecord.CategoryFourteen), data.CategoryFourteen);
            AddTableRow(table, nameof(ExportDataRecord.CategoryFifteen), data.CategoryFifteen);
            AddTableRow(table, nameof(ExportDataRecord.CategorySixteen), data.CategorySixteen);
            AddTableRow(table, nameof(ExportDataRecord.CategorySeventeen), data.CategorySeventeen);
            AddTableRow(table, nameof(ExportDataRecord.CategoryEighteen), data.CategoryEighteen);
            AddTableRow(table, nameof(ExportDataRecord.CategoryNineteen), data.CategoryNineteen);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwenty), data.CategoryTwenty);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyOne), data.CategoryTwentyOne);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyTwo), data.CategoryTwentyTwo);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyThree), data.CategoryTwentyThree);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyFour), data.CategoryTwentyFour);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyFive), data.CategoryTwentyFive);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentySix), data.CategoryTwentySix);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentySeven), data.CategoryTwentySeven);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyEight), data.CategoryTwentyEight);
            AddTableRow(table, nameof(ExportDataRecord.CategoryTwentyNine), data.CategoryTwentyNine);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirty), data.CategoryThirty);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyOne), data.CategoryThirtyOne);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyTwo), data.CategoryThirtyTwo);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyThree), data.CategoryThirtyThree);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyFour), data.CategoryThirtyFour);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyFive), data.CategoryThirtyFive);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtySix), data.CategoryThirtySix);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtySeven), data.CategoryThirtySeven);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyEight), data.CategoryThirtyEight);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyNine), data.CategoryThirtyNine);
            AddTableRow(table, nameof(ExportDataRecord.CategoryThirtyTen), data.CategoryThirtyTen);
            AddTableRow(table, nameof(ExportDataRecord.When), data.When);
            AddTableRow(table, nameof(ExportDataRecord.OriginOne), data.OriginOne);
            AddTableRow(table, nameof(ExportDataRecord.OriginTwo), data.OriginTwo);
            AddTableRow(table, nameof(ExportDataRecord.OriginThree), data.OriginThree);
            AddTableRow(table, nameof(ExportDataRecord.Reference), data.Reference);
            AddTableRow(table, nameof(ExportDataRecord.ContactOne), data.ContactOne);
            AddTableRow(table, nameof(ExportDataRecord.ContactTwo), data.ContactTwo);
            AddTableRow(table, nameof(ExportDataRecord.ContactThree), data.ContactThree);
            AddTableRow(table, nameof(ExportDataRecord.ContactFour), data.ContactFour);
            AddTableRow(table, nameof(ExportDataRecord.ContactFive), data.ContactFive);
            AddTableRow(table, nameof(ExportDataRecord.ContactSix), data.ContactSix);
            AddTableRow(table, nameof(ExportDataRecord.ContactSeven), data.ContactSeven);
            AddTableRow(table, nameof(ExportDataRecord.ContactEight), data.ContactEight);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactOne), data.AlternateContactOne);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactTwo), data.AlternateContactTwo);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactThree), data.AlternateContactThree);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactFour), data.AlternateContactFour);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactFive), data.AlternateContactFive);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactSix), data.AlternateContactSix);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactSeven), data.AlternateContactSeven);
            AddTableRow(table, nameof(ExportDataRecord.AlternateContactEight), data.AlternateContactEight);
            AddTableRow(table, nameof(ExportDataRecord.DeclarationOne), data.DeclarationOne);
            AddTableRow(table, nameof(ExportDataRecord.DeclarationTwo), data.DeclarationTwo);
            AddTableRow(table, nameof(ExportDataRecord.DeclarationThree), data.DeclarationThree);
            AddTableRow(table, nameof(ExportDataRecord.DeclarationFour), data.DeclarationFour);
            AddTableRow(table, nameof(ExportDataRecord.DeclarationFive), data.DeclarationFive);
            AddTableRow(table, nameof(ExportDataRecord.DeclarationSix), data.DeclarationSix);
            AddTableRow(table, nameof(ExportDataRecord.EducationOne), data.EducationOne);
            AddTableRow(table, nameof(ExportDataRecord.EducationTwo), data.EducationTwo);
            AddTableRow(table, nameof(ExportDataRecord.EducationThree), data.EducationThree);
            AddTableRow(table, nameof(ExportDataRecord.EducationFour), data.EducationFour);
            AddTableRow(table, nameof(ExportDataRecord.EducationFive), data.EducationFive);
            AddTableRow(table, nameof(ExportDataRecord.EducationSix), data.EducationSix);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeOne), data.EducationModeOne);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeTwo), data.EducationModeTwo);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeThree), data.EducationModeThree);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeFour), data.EducationModeFour);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeFive), data.EducationModeFive);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeSix), data.EducationModeSix);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeSeven), data.EducationModeSeven);
            AddTableRow(table, nameof(ExportDataRecord.EducationModeEight), data.EducationModeEight);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactSeven), data.ExternalContactSeven);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactEight), data.ExternalContactEight);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactNine), data.ExternalContactNine);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTen), data.ExternalContactTen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactEleven), data.ExternalContactEleven);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwelve), data.ExternalContactTwelve);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactThirteen), data.ExternalContactThirteen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactFourteen), data.ExternalContactFourteen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactFifteen), data.ExternalContactFifteen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactSixteen), data.ExternalContactSixteen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactSeventeen), data.ExternalContactSeventeen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactEighteen), data.ExternalContactEighteen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactNineteen), data.ExternalContactNineteen);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwenty), data.ExternalContactTwenty);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentyOne), data.ExternalContactTwentyOne);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentyTwo), data.ExternalContactTwentyTwo);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentyThree), data.ExternalContactTwentyThree);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentyFour), data.ExternalContactTwentyFour);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentyFive), data.ExternalContactTwentyFive);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentySix), data.ExternalContactTwentySix);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentySeven), data.ExternalContactTwentySeven);
            AddTableRow(table, nameof(ExportDataRecord.ExternalContactTwentyEight), data.ExternalContactTwentyEight);

            docBody.Append(table);
        }

        public void AddTableRow(Table documentTable, string label, string value)
        {
            var row = new TableRow();

            var labelCell = new TableCell();
            labelCell.Append(GetCellProperties());
            var labelParagraph = new Paragraph(new Run(new Text(label)));
            labelCell.Append(labelParagraph);
            row.Append(labelCell);

            var valueCell = new TableCell();
            valueCell.Append(GetCellProperties());
            var valueParagraph = new Paragraph(new Run(new Text(value)));
            valueCell.Append(valueParagraph);
            row.Append(valueCell);

            documentTable.Append(row);
        }

        public void AddTableRow(Table documentTable, string label, bool value)
        {
            AddTableRow(documentTable, label, value.ToYesNo());
        }

        public void AddTableRow(Table documentTable, string label, DateTime value)
        {
            AddTableRow(documentTable, label, value.ToString("dd/MM/yyyy"));
        }

        public void AddTableRow(Table documentTable, string label, decimal value)
        {
            AddTableRow(documentTable, label, value.ToString("F2"));
        }

        public TableCellProperties GetCellProperties()
        {
            return new TableCellProperties
                       {
                           TableCellMargin = new TableCellMargin
                                                 {
                                                     TopMargin = new TopMargin { Width = "20", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) },
                                                     BottomMargin = new BottomMargin { Width = "20", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) },
                                                     LeftMargin = new LeftMargin { Width = "20", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) },
                                                     RightMargin = new RightMargin { Width = "20", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) }
                                                 }
                       };
        }
    }
}