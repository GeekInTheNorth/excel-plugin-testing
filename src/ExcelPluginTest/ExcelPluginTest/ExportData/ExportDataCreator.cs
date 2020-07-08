namespace ExcelPluginTest.ExportData
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public class ExportDataCreator
    {
        private readonly Random random;

        public ExportDataCreator()
        {
            this.random = new Random();
        }

        public List<ExportDataRecord> Create(int records)
        {
            return Enumerable.Range(0, records).Select(x => CreateRecord()).ToList();
        }

        private ExportDataRecord CreateRecord()
        {
            return new ExportDataRecord
                       {
                           Title = RandomString(10),
                           Forename = RandomString(10),
                           Surname = RandomString(10),
                           PreviousSurname = RandomString(10),
                           Profession = RandomString(10),
                           Number = 1234.56M,
                           NumberReference = RandomString(10),
                           NumberDate = RandomDate(),
                           PreviouslyProvided = RandomBool(),
                           PreviouslyProvidedReference = RandomString(10),
                           CategoryOne = RandomBool(),
                           CategoryTwo = RandomBool(),
                           CategoryThree = RandomBool(),
                           CategoryFour = RandomBool(),
                           CategoryFive = RandomBool(),
                           CategorySix = RandomBool(),
                           CategorySeven = RandomBool(),
                           CategoryEight = RandomBool(),
                           CategoryNine = RandomBool(),
                           CategoryTen = RandomBool(),
                           CategoryEleven = RandomBool(),
                           CategoryTwelve = RandomBool(),
                           CategoryThirteen = RandomBool(),
                           CategoryFourteen = RandomBool(),
                           CategoryFifteen = RandomBool(),
                           CategorySixteen = RandomBool(),
                           CategorySeventeen = RandomBool(),
                           CategoryEighteen = RandomBool(),
                           CategoryNineteen = RandomBool(),
                           CategoryTwenty = RandomBool(),
                           CategoryTwentyOne = RandomBool(),
                           CategoryTwentyTwo = RandomBool(),
                           CategoryTwentyThree = RandomBool(),
                           CategoryTwentyFour = RandomBool(),
                           CategoryTwentyFive = RandomBool(),
                           CategoryTwentySix = RandomBool(),
                           CategoryTwentySeven = RandomBool(),
                           CategoryTwentyEight = RandomBool(),
                           CategoryTwentyNine = RandomBool(),
                           CategoryThirty = RandomBool(),
                           CategoryThirtyOne = RandomBool(),
                           CategoryThirtyTwo = RandomBool(),
                           CategoryThirtyThree = RandomBool(),
                           CategoryThirtyFour = RandomBool(),
                           CategoryThirtyFive = RandomBool(),
                           CategoryThirtySix = RandomBool(),
                           CategoryThirtySeven = RandomBool(),
                           CategoryThirtyEight = RandomBool(),
                           CategoryThirtyNine = RandomBool(),
                           CategoryThirtyTen = RandomBool(),
                           When = RandomDate(),
                           OriginOne = RandomString(10),
                           OriginTwo = RandomString(10),
                           OriginThree = RandomString(10),
                           Reference = RandomString(10),
                           ContactOne = RandomString(10),
                           ContactTwo = RandomString(10),
                           ContactThree = RandomString(10),
                           ContactFour = RandomString(10),
                           ContactFive = RandomString(10),
                           ContactSix = RandomString(10),
                           ContactSeven = RandomString(10),
                           ContactEight = RandomString(10),
                           AlternateContactOne = RandomString(10),
                           AlternateContactTwo = RandomString(10),
                           AlternateContactThree = RandomString(10),
                           AlternateContactFour = RandomString(10),
                           AlternateContactFive = RandomString(10),
                           AlternateContactSix = RandomString(10),
                           AlternateContactSeven = RandomString(10),
                           AlternateContactEight = RandomString(10),
                           DeclarationOne = RandomBool(),
                           DeclarationTwo = RandomBool(),
                           DeclarationThree = RandomBool(),
                           DeclarationFour = RandomBool(),
                           DeclarationFive = RandomBool(),
                           DeclarationSix = RandomBool(),
                           EducationOne = RandomString(10),
                           EducationTwo = RandomDate(),
                           EducationThree = RandomDate(),
                           EducationFour = RandomString(10),
                           EducationFive = RandomString(10),
                           EducationSix = RandomString(10),
                           EducationModeOne = RandomBool(),
                           EducationModeTwo = RandomBool(),
                           EducationModeThree = RandomBool(),
                           EducationModeFour = RandomBool(),
                           EducationModeFive = RandomBool(),
                           EducationModeSix = RandomBool(),
                           EducationModeSeven = RandomBool(),
                           EducationModeEight = RandomBool(),
                           ExternalContactSeven = RandomString(10),
                           ExternalContactEight = RandomString(10),
                           ExternalContactNine = RandomString(10),
                           ExternalContactTen = RandomString(10),
                           ExternalContactEleven = RandomString(10),
                           ExternalContactTwelve = RandomString(10),
                           ExternalContactThirteen = RandomString(10),
                           ExternalContactFourteen = RandomString(10),
                           ExternalContactFifteen = RandomString(10),
                           ExternalContactSixteen = RandomString(10),
                           ExternalContactSeventeen = RandomString(10),
                           ExternalContactEighteen = RandomString(10),
                           ExternalContactNineteen = RandomString(10),
                           ExternalContactTwenty = RandomDate(),
                           ExternalContactTwentyOne = RandomDate(),
                           ExternalContactTwentyTwo = RandomString(10),
                           ExternalContactTwentyThree = RandomString(10),
                           ExternalContactTwentyFour = RandomString(10),
                           ExternalContactTwentyFive = RandomString(10),
                           ExternalContactTwentySix = RandomString(10),
                           ExternalContactTwentySeven = RandomString(10),
                           ExternalContactTwentyEight = RandomBool()
                       };
        }

        private string RandomString(int length)
        {
            const string chars = @"ABCDEFGHIJKLMNOPQRSTUVWXYZ ";

            return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private bool RandomBool()
        {
            return random.Next(2) == 1;
        }

        private DateTime RandomDate()
        {
            return DateTime.Today.AddDays(random.Next(-1000, 1000));
        }
    }
}