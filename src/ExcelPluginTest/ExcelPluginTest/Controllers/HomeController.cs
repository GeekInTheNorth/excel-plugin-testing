namespace ExcelPluginTest.Controllers
{
    using System.Web.Mvc;

    using ExcelPluginTest.ClosedXml;
    using ExcelPluginTest.Interfaces;
    using ExcelPluginTest.Models;

    public class HomeController : Controller
    {
        private IExcelCreator excelCreator;

        public HomeController()
        {
            excelCreator = new ClosedXmlCreator();
        }

        public ActionResult Index()
        {
            var model = new HomeViewModel();

            return this.View(model);
        }

        public FileResult Document()
        {
            var document = excelCreator.Create();

            return this.File(document, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
    }
}