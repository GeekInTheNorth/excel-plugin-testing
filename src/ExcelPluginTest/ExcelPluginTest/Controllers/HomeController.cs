namespace ExcelPluginTest.Controllers
{
    using System.Web.Mvc;

    using ExcelPluginTest.ClosedXml;
    using ExcelPluginTest.Interfaces;
    using ExcelPluginTest.Models;

    public class HomeController : Controller
    {
        private IExcelCreator excelCreator;

        private IWordCreator wordCreator;

        public HomeController()
        {
            excelCreator = new ClosedXmlExcelCreator();
            wordCreator = new ClosedXmlWordCreator();
        }

        public ActionResult Index()
        {
            var model = new HomeViewModel();

            return this.View(model);
        }

        public FileResult Excel()
        {
            var document = excelCreator.Create();

            return this.File(document, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "excel-test.xlsx");
        }

        public FileResult Document()
        {
            var document = wordCreator.Create();

            return this.File(document, "application/vnd.openxmlformats-officedocument.wordprocessing", "word-test.docx");
        }
    }
}