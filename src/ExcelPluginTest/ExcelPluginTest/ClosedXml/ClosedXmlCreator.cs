// --------------------------------------------------------------------------------------------------------------------
// <copyright company="twentysix" file="ClosedXmlCreator.cs">
// Copyright (c) twentysix.  All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace ExcelPluginTest.ClosedXml
{
    using System.IO;

    using ClosedXML.Excel;

    using ExcelPluginTest.Interfaces;

    public class ClosedXmlCreator : IExcelCreator
    {
        public byte[] Create()
        {
            using (var stream = new MemoryStream())
            {
                var workBook = new XLWorkbook();
                var workSheet = workBook.Worksheets.Add("Test Sheet");
                workSheet.Cell("A1").Value = "Hello World";

                workBook.SaveAs(stream);

                return stream.ToArray();
            }
        }
    }
}