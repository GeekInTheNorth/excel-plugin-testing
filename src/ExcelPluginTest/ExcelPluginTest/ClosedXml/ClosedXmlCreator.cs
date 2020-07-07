namespace ExcelPluginTest.ClosedXml
{
    using System.IO;

    using ClosedXML.Excel;

    using ExcelPluginTest.ExportData;
    using ExcelPluginTest.Interfaces;

    public class ClosedXmlCreator : IExcelCreator
    {
        private readonly ExportDataCreator _creator;

        public ClosedXmlCreator()
        {
            _creator = new ExportDataCreator();
        }

        public byte[] Create()
        {
            var data = _creator.Create(1000);

            using (var stream = new MemoryStream())
            {
                var workBook = new XLWorkbook();
                var workSheet = workBook.Worksheets.Add("Test Sheet");

                // Header Row
                workSheet.Cell(1, 1).Value = nameof(ExportDataRecord.Title);
                workSheet.Cell(1, 2).Value = nameof(ExportDataRecord.Forename);
                workSheet.Cell(1, 3).Value = nameof(ExportDataRecord.Surname);
                workSheet.Cell(1, 4).Value = nameof(ExportDataRecord.PreviousSurname);
                workSheet.Cell(1, 5).Value = nameof(ExportDataRecord.Profession);
                workSheet.Cell(1, 6).Value = nameof(ExportDataRecord.Number);
                workSheet.Cell(1, 7).Value = nameof(ExportDataRecord.NumberReference);
                workSheet.Cell(1, 8).Value = nameof(ExportDataRecord.NumberDate);
                workSheet.Cell(1, 9).Value = nameof(ExportDataRecord.PreviouslyProvided);
                workSheet.Cell(1, 10).Value = nameof(ExportDataRecord.PreviouslyProvidedReference);
                workSheet.Cell(1, 11).Value = nameof(ExportDataRecord.CategoryOne);
                workSheet.Cell(1, 12).Value = nameof(ExportDataRecord.CategoryTwo);
                workSheet.Cell(1, 13).Value = nameof(ExportDataRecord.CategoryThree);
                workSheet.Cell(1, 14).Value = nameof(ExportDataRecord.CategoryFour);
                workSheet.Cell(1, 15).Value = nameof(ExportDataRecord.CategoryFive);
                workSheet.Cell(1, 16).Value = nameof(ExportDataRecord.CategorySix);
                workSheet.Cell(1, 17).Value = nameof(ExportDataRecord.CategorySeven);
                workSheet.Cell(1, 18).Value = nameof(ExportDataRecord.CategoryEight);
                workSheet.Cell(1, 19).Value = nameof(ExportDataRecord.CategoryNine);
                workSheet.Cell(1, 20).Value = nameof(ExportDataRecord.CategoryTen);
                workSheet.Cell(1, 21).Value = nameof(ExportDataRecord.CategoryEleven);
                workSheet.Cell(1, 22).Value = nameof(ExportDataRecord.CategoryTwelve);
                workSheet.Cell(1, 23).Value = nameof(ExportDataRecord.CategoryThirteen);
                workSheet.Cell(1, 24).Value = nameof(ExportDataRecord.CategoryFourteen);
                workSheet.Cell(1, 25).Value = nameof(ExportDataRecord.CategoryFifteen);
                workSheet.Cell(1, 26).Value = nameof(ExportDataRecord.CategorySixteen);
                workSheet.Cell(1, 27).Value = nameof(ExportDataRecord.CategorySeventeen);
                workSheet.Cell(1, 28).Value = nameof(ExportDataRecord.CategoryEighteen);
                workSheet.Cell(1, 29).Value = nameof(ExportDataRecord.CategoryNineteen);
                workSheet.Cell(1, 30).Value = nameof(ExportDataRecord.CategoryTwenty);
                workSheet.Cell(1, 31).Value = nameof(ExportDataRecord.CategoryTwentyOne);
                workSheet.Cell(1, 32).Value = nameof(ExportDataRecord.CategoryTwentyTwo);
                workSheet.Cell(1, 33).Value = nameof(ExportDataRecord.CategoryTwentyThree);
                workSheet.Cell(1, 34).Value = nameof(ExportDataRecord.CategoryTwentyFour);
                workSheet.Cell(1, 35).Value = nameof(ExportDataRecord.CategoryTwentyFive);
                workSheet.Cell(1, 36).Value = nameof(ExportDataRecord.CategoryTwentySix);
                workSheet.Cell(1, 37).Value = nameof(ExportDataRecord.CategoryTwentySeven);
                workSheet.Cell(1, 38).Value = nameof(ExportDataRecord.CategoryTwentyEight);
                workSheet.Cell(1, 39).Value = nameof(ExportDataRecord.CategoryTwentyNine);
                workSheet.Cell(1, 40).Value = nameof(ExportDataRecord.CategoryThirty);
                workSheet.Cell(1, 41).Value = nameof(ExportDataRecord.CategoryThirtyOne);
                workSheet.Cell(1, 42).Value = nameof(ExportDataRecord.CategoryThirtyTwo);
                workSheet.Cell(1, 43).Value = nameof(ExportDataRecord.CategoryThirtyThree);
                workSheet.Cell(1, 44).Value = nameof(ExportDataRecord.CategoryThirtyFour);
                workSheet.Cell(1, 45).Value = nameof(ExportDataRecord.CategoryThirtyFive);
                workSheet.Cell(1, 46).Value = nameof(ExportDataRecord.CategoryThirtySix);
                workSheet.Cell(1, 47).Value = nameof(ExportDataRecord.CategoryThirtySeven);
                workSheet.Cell(1, 48).Value = nameof(ExportDataRecord.CategoryThirtyEight);
                workSheet.Cell(1, 49).Value = nameof(ExportDataRecord.CategoryThirtyNine);
                workSheet.Cell(1, 50).Value = nameof(ExportDataRecord.CategoryThirtyTen);
                workSheet.Cell(1, 51).Value = nameof(ExportDataRecord.When);
                workSheet.Cell(1, 52).Value = nameof(ExportDataRecord.OriginOne);
                workSheet.Cell(1, 53).Value = nameof(ExportDataRecord.OriginTwo);
                workSheet.Cell(1, 54).Value = nameof(ExportDataRecord.OriginThree);
                workSheet.Cell(1, 55).Value = nameof(ExportDataRecord.Reference);
                workSheet.Cell(1, 56).Value = nameof(ExportDataRecord.ContactOne);
                workSheet.Cell(1, 57).Value = nameof(ExportDataRecord.ContactTwo);
                workSheet.Cell(1, 58).Value = nameof(ExportDataRecord.ContactThree);
                workSheet.Cell(1, 59).Value = nameof(ExportDataRecord.ContactFour);
                workSheet.Cell(1, 60).Value = nameof(ExportDataRecord.ContactFive);
                workSheet.Cell(1, 61).Value = nameof(ExportDataRecord.ContactSix);
                workSheet.Cell(1, 62).Value = nameof(ExportDataRecord.ContactSeven);
                workSheet.Cell(1, 63).Value = nameof(ExportDataRecord.ContactEight);
                workSheet.Cell(1, 64).Value = nameof(ExportDataRecord.AlternateContactOne);
                workSheet.Cell(1, 65).Value = nameof(ExportDataRecord.AlternateContactTwo);
                workSheet.Cell(1, 66).Value = nameof(ExportDataRecord.AlternateContactThree);
                workSheet.Cell(1, 67).Value = nameof(ExportDataRecord.AlternateContactFour);
                workSheet.Cell(1, 68).Value = nameof(ExportDataRecord.AlternateContactFive);
                workSheet.Cell(1, 69).Value = nameof(ExportDataRecord.AlternateContactSix);
                workSheet.Cell(1, 70).Value = nameof(ExportDataRecord.AlternateContactSeven);
                workSheet.Cell(1, 71).Value = nameof(ExportDataRecord.AlternateContactEight);
                workSheet.Cell(1, 72).Value = nameof(ExportDataRecord.DeclarationOne);
                workSheet.Cell(1, 73).Value = nameof(ExportDataRecord.DeclarationTwo);
                workSheet.Cell(1, 74).Value = nameof(ExportDataRecord.DeclarationThree);
                workSheet.Cell(1, 75).Value = nameof(ExportDataRecord.DeclarationFour);
                workSheet.Cell(1, 76).Value = nameof(ExportDataRecord.DeclarationFive);
                workSheet.Cell(1, 77).Value = nameof(ExportDataRecord.DeclarationSix);
                workSheet.Cell(1, 78).Value = nameof(ExportDataRecord.EducationOne);
                workSheet.Cell(1, 79).Value = nameof(ExportDataRecord.EducationTwo);
                workSheet.Cell(1, 80).Value = nameof(ExportDataRecord.EducationThree);
                workSheet.Cell(1, 81).Value = nameof(ExportDataRecord.EducationFour);
                workSheet.Cell(1, 82).Value = nameof(ExportDataRecord.EducationFive);
                workSheet.Cell(1, 83).Value = nameof(ExportDataRecord.EducationSix);
                workSheet.Cell(1, 84).Value = nameof(ExportDataRecord.EducationModeOne);
                workSheet.Cell(1, 85).Value = nameof(ExportDataRecord.EducationModeTwo);
                workSheet.Cell(1, 86).Value = nameof(ExportDataRecord.EducationModeThree);
                workSheet.Cell(1, 87).Value = nameof(ExportDataRecord.EducationModeFour);
                workSheet.Cell(1, 88).Value = nameof(ExportDataRecord.EducationModeFive);
                workSheet.Cell(1, 89).Value = nameof(ExportDataRecord.EducationModeSix);
                workSheet.Cell(1, 90).Value = nameof(ExportDataRecord.EducationModeSeven);
                workSheet.Cell(1, 91).Value = nameof(ExportDataRecord.EducationModeEight);
                workSheet.Cell(1, 92).Value = nameof(ExportDataRecord.ExternalContactSeven);
                workSheet.Cell(1, 93).Value = nameof(ExportDataRecord.ExternalContactEight);
                workSheet.Cell(1, 94).Value = nameof(ExportDataRecord.ExternalContactNine);
                workSheet.Cell(1, 95).Value = nameof(ExportDataRecord.ExternalContactTen);
                workSheet.Cell(1, 96).Value = nameof(ExportDataRecord.ExternalContactEleven);
                workSheet.Cell(1, 97).Value = nameof(ExportDataRecord.ExternalContactTwelve);
                workSheet.Cell(1, 98).Value = nameof(ExportDataRecord.ExternalContactThirteen);
                workSheet.Cell(1, 99).Value = nameof(ExportDataRecord.ExternalContactFourteen);
                workSheet.Cell(1, 100).Value = nameof(ExportDataRecord.ExternalContactFifteen);
                workSheet.Cell(1, 101).Value = nameof(ExportDataRecord.ExternalContactSixteen);
                workSheet.Cell(1, 102).Value = nameof(ExportDataRecord.ExternalContactSeventeen);
                workSheet.Cell(1, 103).Value = nameof(ExportDataRecord.ExternalContactEighteen);
                workSheet.Cell(1, 104).Value = nameof(ExportDataRecord.ExternalContactNineteen);
                workSheet.Cell(1, 105).Value = nameof(ExportDataRecord.ExternalContactTwenty);
                workSheet.Cell(1, 106).Value = nameof(ExportDataRecord.ExternalContactTwentyOne);
                workSheet.Cell(1, 107).Value = nameof(ExportDataRecord.ExternalContactTwentyTwo);
                workSheet.Cell(1, 108).Value = nameof(ExportDataRecord.ExternalContactTwentyThree);
                workSheet.Cell(1, 109).Value = nameof(ExportDataRecord.ExternalContactTwentyFour);
                workSheet.Cell(1, 110).Value = nameof(ExportDataRecord.ExternalContactTwentyFive);
                workSheet.Cell(1, 111).Value = nameof(ExportDataRecord.ExternalContactTwentySix);
                workSheet.Cell(1, 112).Value = nameof(ExportDataRecord.ExternalContactTwentySeven);
                workSheet.Cell(1, 113).Value = nameof(ExportDataRecord.ExternalContactTwentyEight);

                var row = 2;
                foreach (var dataRecord in data)
                {
                    workSheet.Cell(row, 1).Value = dataRecord.Title;
                    workSheet.Cell(row, 2).Value = dataRecord.Forename;
                    workSheet.Cell(row, 3).Value = dataRecord.Surname;
                    workSheet.Cell(row, 4).Value = dataRecord.PreviousSurname;
                    workSheet.Cell(row, 5).Value = dataRecord.Profession;
                    workSheet.Cell(row, 6).Value = dataRecord.Number;
                    workSheet.Cell(row, 7).Value = dataRecord.NumberReference;
                    workSheet.Cell(row, 8).Value = dataRecord.NumberDate;
                    workSheet.Cell(row, 9).Value = dataRecord.PreviouslyProvided;
                    workSheet.Cell(row, 10).Value = dataRecord.PreviouslyProvidedReference;
                    workSheet.Cell(row, 11).Value = dataRecord.CategoryOne;
                    workSheet.Cell(row, 12).Value = dataRecord.CategoryTwo;
                    workSheet.Cell(row, 13).Value = dataRecord.CategoryThree;
                    workSheet.Cell(row, 14).Value = dataRecord.CategoryFour;
                    workSheet.Cell(row, 15).Value = dataRecord.CategoryFive;
                    workSheet.Cell(row, 16).Value = dataRecord.CategorySix;
                    workSheet.Cell(row, 17).Value = dataRecord.CategorySeven;
                    workSheet.Cell(row, 18).Value = dataRecord.CategoryEight;
                    workSheet.Cell(row, 19).Value = dataRecord.CategoryNine;
                    workSheet.Cell(row, 20).Value = dataRecord.CategoryTen;
                    workSheet.Cell(row, 21).Value = dataRecord.CategoryEleven;
                    workSheet.Cell(row, 22).Value = dataRecord.CategoryTwelve;
                    workSheet.Cell(row, 23).Value = dataRecord.CategoryThirteen;
                    workSheet.Cell(row, 24).Value = dataRecord.CategoryFourteen;
                    workSheet.Cell(row, 25).Value = dataRecord.CategoryFifteen;
                    workSheet.Cell(row, 26).Value = dataRecord.CategorySixteen;
                    workSheet.Cell(row, 27).Value = dataRecord.CategorySeventeen;
                    workSheet.Cell(row, 28).Value = dataRecord.CategoryEighteen;
                    workSheet.Cell(row, 29).Value = dataRecord.CategoryNineteen;
                    workSheet.Cell(row, 30).Value = dataRecord.CategoryTwenty;
                    workSheet.Cell(row, 31).Value = dataRecord.CategoryTwentyOne;
                    workSheet.Cell(row, 32).Value = dataRecord.CategoryTwentyTwo;
                    workSheet.Cell(row, 33).Value = dataRecord.CategoryTwentyThree;
                    workSheet.Cell(row, 34).Value = dataRecord.CategoryTwentyFour;
                    workSheet.Cell(row, 35).Value = dataRecord.CategoryTwentyFive;
                    workSheet.Cell(row, 36).Value = dataRecord.CategoryTwentySix;
                    workSheet.Cell(row, 37).Value = dataRecord.CategoryTwentySeven;
                    workSheet.Cell(row, 38).Value = dataRecord.CategoryTwentyEight;
                    workSheet.Cell(row, 39).Value = dataRecord.CategoryTwentyNine;
                    workSheet.Cell(row, 40).Value = dataRecord.CategoryThirty;
                    workSheet.Cell(row, 41).Value = dataRecord.CategoryThirtyOne;
                    workSheet.Cell(row, 42).Value = dataRecord.CategoryThirtyTwo;
                    workSheet.Cell(row, 43).Value = dataRecord.CategoryThirtyThree;
                    workSheet.Cell(row, 44).Value = dataRecord.CategoryThirtyFour;
                    workSheet.Cell(row, 45).Value = dataRecord.CategoryThirtyFive;
                    workSheet.Cell(row, 46).Value = dataRecord.CategoryThirtySix;
                    workSheet.Cell(row, 47).Value = dataRecord.CategoryThirtySeven;
                    workSheet.Cell(row, 48).Value = dataRecord.CategoryThirtyEight;
                    workSheet.Cell(row, 49).Value = dataRecord.CategoryThirtyNine;
                    workSheet.Cell(row, 50).Value = dataRecord.CategoryThirtyTen;
                    workSheet.Cell(row, 51).Value = dataRecord.When;
                    workSheet.Cell(row, 52).Value = dataRecord.OriginOne;
                    workSheet.Cell(row, 53).Value = dataRecord.OriginTwo;
                    workSheet.Cell(row, 54).Value = dataRecord.OriginThree;
                    workSheet.Cell(row, 55).Value = dataRecord.Reference;
                    workSheet.Cell(row, 56).Value = dataRecord.ContactOne;
                    workSheet.Cell(row, 57).Value = dataRecord.ContactTwo;
                    workSheet.Cell(row, 58).Value = dataRecord.ContactThree;
                    workSheet.Cell(row, 59).Value = dataRecord.ContactFour;
                    workSheet.Cell(row, 60).Value = dataRecord.ContactFive;
                    workSheet.Cell(row, 61).Value = dataRecord.ContactSix;
                    workSheet.Cell(row, 62).Value = dataRecord.ContactSeven;
                    workSheet.Cell(row, 63).Value = dataRecord.ContactEight;
                    workSheet.Cell(row, 64).Value = dataRecord.AlternateContactOne;
                    workSheet.Cell(row, 65).Value = dataRecord.AlternateContactTwo;
                    workSheet.Cell(row, 66).Value = dataRecord.AlternateContactThree;
                    workSheet.Cell(row, 67).Value = dataRecord.AlternateContactFour;
                    workSheet.Cell(row, 68).Value = dataRecord.AlternateContactFive;
                    workSheet.Cell(row, 69).Value = dataRecord.AlternateContactSix;
                    workSheet.Cell(row, 70).Value = dataRecord.AlternateContactSeven;
                    workSheet.Cell(row, 71).Value = dataRecord.AlternateContactEight;
                    workSheet.Cell(row, 72).Value = dataRecord.DeclarationOne;
                    workSheet.Cell(row, 73).Value = dataRecord.DeclarationTwo;
                    workSheet.Cell(row, 74).Value = dataRecord.DeclarationThree;
                    workSheet.Cell(row, 75).Value = dataRecord.DeclarationFour;
                    workSheet.Cell(row, 76).Value = dataRecord.DeclarationFive;
                    workSheet.Cell(row, 77).Value = dataRecord.DeclarationSix;
                    workSheet.Cell(row, 78).Value = dataRecord.EducationOne;
                    workSheet.Cell(row, 79).Value = dataRecord.EducationTwo;
                    workSheet.Cell(row, 80).Value = dataRecord.EducationThree;
                    workSheet.Cell(row, 81).Value = dataRecord.EducationFour;
                    workSheet.Cell(row, 82).Value = dataRecord.EducationFive;
                    workSheet.Cell(row, 83).Value = dataRecord.EducationSix;
                    workSheet.Cell(row, 84).Value = dataRecord.EducationModeOne;
                    workSheet.Cell(row, 85).Value = dataRecord.EducationModeTwo;
                    workSheet.Cell(row, 86).Value = dataRecord.EducationModeThree;
                    workSheet.Cell(row, 87).Value = dataRecord.EducationModeFour;
                    workSheet.Cell(row, 88).Value = dataRecord.EducationModeFive;
                    workSheet.Cell(row, 89).Value = dataRecord.EducationModeSix;
                    workSheet.Cell(row, 90).Value = dataRecord.EducationModeSeven;
                    workSheet.Cell(row, 91).Value = dataRecord.EducationModeEight;
                    workSheet.Cell(row, 92).Value = dataRecord.ExternalContactSeven;
                    workSheet.Cell(row, 93).Value = dataRecord.ExternalContactEight;
                    workSheet.Cell(row, 94).Value = dataRecord.ExternalContactNine;
                    workSheet.Cell(row, 95).Value = dataRecord.ExternalContactTen;
                    workSheet.Cell(row, 96).Value = dataRecord.ExternalContactEleven;
                    workSheet.Cell(row, 97).Value = dataRecord.ExternalContactTwelve;
                    workSheet.Cell(row, 98).Value = dataRecord.ExternalContactThirteen;
                    workSheet.Cell(row, 99).Value = dataRecord.ExternalContactFourteen;
                    workSheet.Cell(row, 100).Value = dataRecord.ExternalContactFifteen;
                    workSheet.Cell(row, 101).Value = dataRecord.ExternalContactSixteen;
                    workSheet.Cell(row, 102).Value = dataRecord.ExternalContactSeventeen;
                    workSheet.Cell(row, 103).Value = dataRecord.ExternalContactEighteen;
                    workSheet.Cell(row, 104).Value = dataRecord.ExternalContactNineteen;
                    workSheet.Cell(row, 105).Value = dataRecord.ExternalContactTwenty;
                    workSheet.Cell(row, 106).Value = dataRecord.ExternalContactTwentyOne;
                    workSheet.Cell(row, 107).Value = dataRecord.ExternalContactTwentyTwo;
                    workSheet.Cell(row, 108).Value = dataRecord.ExternalContactTwentyThree;
                    workSheet.Cell(row, 109).Value = dataRecord.ExternalContactTwentyFour;
                    workSheet.Cell(row, 110).Value = dataRecord.ExternalContactTwentyFive;
                    workSheet.Cell(row, 111).Value = dataRecord.ExternalContactTwentySix;
                    workSheet.Cell(row, 112).Value = dataRecord.ExternalContactTwentySeven;
                    workSheet.Cell(row, 113).Value = dataRecord.ExternalContactTwentyEight;

                    row++;
                }

                workBook.SaveAs(stream);

                return stream.ToArray();
            }
        }
    }
}