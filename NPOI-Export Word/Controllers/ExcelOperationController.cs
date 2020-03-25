
/**
 * Author:追逐时光
 * Description:Excel导入导出操作
 */
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using YY_Services;
using YY_Utility;

namespace YY_NPOI_ExportAndImport.Controllers
{
    public class ExcelOperationController : Controller
    {
        private readonly NpoiExcelOperationService _excelExport;

        public ExcelOperationController(NpoiExcelOperationService excelExport)
        {
            _excelExport = excelExport;
        }

        /// <summary>
        /// 导出Excel文档展示界面
        /// </summary>
        /// <returns></returns>
        public IActionResult ExportExcelData()
        {
            return View();
        }

        /// <summary>
        /// Excel文档生成并导出
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public JsonResult DataExportSimple()
        {
            bool result = _excelExport.ExcelDataExport(out string resultMsg, out string excelFilePath);

            return Json(result == true ? new { code = 1, data = excelFilePath } : new { code = 0, data = resultMsg });
        }


    }
}