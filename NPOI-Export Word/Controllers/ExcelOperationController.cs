
/**
 * Author:追逐时光
 * Description:Excel导入导出操作
 */

using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using YY_Model;
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

        public ActionResult Index()
        {
            return View();
        }

        ///// <summary>
        ///// 获取用户信息
        ///// </summary>
        ///// <param name="page"></param>
        ///// <param name="limit"></param>
        ///// <returns></returns>
        //[HttpGet]
        //public JsonResult GetUserInfo(int page = 1, int limit = 15)
        //{
        //    try
        //    {
        //        //使用ef--skip().take()进行数据分页前面必须增加orderby，否则报错
        //        var List = UserEntites.UserInfo.OrderBy(p => p.Id).Skip((page - 1) * limit).Take(limit).ToList();

        //        return Json(new { code = 0, count = UserEntites.UserInfo.Count(), data = List }, JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new { code = 1, msg = ex.Message });
        //    }}


        /// <summary>
        /// Excel数据导入
        /// </summary>
        /// <param name="FileStram"></param>
        /// <returns></returns>
        //public ActionResult DataImport(HttpPostedFileBase file)
        //{
        //    var message = "";
        //    int Columns = 0;
        //    //判断是否提交excel文件
        //    var FileName = file.FileName.Split('.');
        //    if (file != null && file.ContentLength > 0)
        //    {
        //        if (FileName[1] == "xls" || FileName[1] == "xlsx")
        //        {
        //            //首先我们需要导入数据的话第一步其实就是先把excel数据保存到本地中，然后通过Npoi封装的方法去读取已保存的Excel数据

        //            string DictorysPath = Server.MapPath("~/Content/ExcelFiles/" + DateTime.Now.ToString("yyyyMMdd"));
        //            if (!System.IO.Directory.Exists(DictorysPath))
        //            {
        //                System.IO.Directory.CreateDirectory(DictorysPath);
        //            }

        //            file.SaveAs(System.IO.Path.Combine(DictorysPath, file.FileName));

        //            //将Excel数据转化为DataTable数据源
        //            DataTable Dt = NpoiHelper.Import(System.IO.Path.Combine(DictorysPath, file.FileName), FileName[1]);
        //            List<UserInfo> list = new List<UserInfo>();

        //            for (int i = 0; i < Dt.Rows.Count; i++)
        //            {
        //                UserInfo U = new UserInfo();
        //                //从行索引从1开始，标题除外
        //                U.UserName = Dt.Rows[i][0].ToString();
        //                U.Sex = Dt.Rows[i][1].ToString();
        //                U.Phone = Dt.Rows[i][2].ToString();
        //                U.Hobby = Dt.Rows[i][3].ToString();
        //                list.Add(U);
        //            }

        //            //数据全部添加
        //            UserEntites.Set<UserInfo>().AddRange(list);
        //            Columns = UserEntites.SaveChanges();
        //            if (Columns > 0)
        //            {
        //                message = "导入成功";
        //            }
        //            else
        //            {
        //                message = "导入失败";
        //            }

        //        }
        //        else
        //        {
        //            message = "格式错误";
        //        }
        //    }
        //    else
        //    {
        //        message = "未找到需要导入的数据";
        //    }
        //    ViewBag.Columns = Columns;
        //    ViewBag.Message = message;
        //    return View();
        //}


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