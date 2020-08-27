using System.Diagnostics;
using Export.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using YY_NpoiExportAndImport.Models;

namespace YY_NpoiExportAndImport.Controllers
{
    public class HomeController : Controller
    {
        private readonly NpoiWordExportService _exportService;

        /// <summary>
        /// 构造函数依赖注入
        /// </summary>
        /// <param name="noExportService"></param>
        public HomeController(NpoiWordExportService noExportService)
        {
            _exportService = noExportService;
        }

        /// <summary>
        /// Word Export
        /// </summary>
        /// <returns></returns>
        public JsonResult WordExport()
        {
            bool result = _exportService.SaveWordFile(out string savePath);

            return Json(result == true ? new { code = 1, data = savePath } : new { code = 0, data = savePath });
        }

        public IActionResult Index()
        {
            return View();
        }


        public IActionResult Privacy()
        {
            return View();
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
