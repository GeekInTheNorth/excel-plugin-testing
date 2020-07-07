namespace ExcelPluginTest.Controllers
{
    using System.Web.Mvc;

    using ExcelPluginTest.Models;

    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            var model = new HomeViewModel();

            return View(model);
        }
    }
}