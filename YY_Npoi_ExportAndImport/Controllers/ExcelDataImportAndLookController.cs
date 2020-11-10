using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.AspNetCore.Http;
//using YY_Dal;
using YY_Model;
using YY_Services;
using ActionResult = Microsoft.AspNetCore.Mvc.ActionResult;
using Controller = Microsoft.AspNetCore.Mvc.Controller;
using JsonResult = Microsoft.AspNetCore.Mvc.JsonResult;

namespace YY_NpoiExportAndImport.Controllers
{
    public class ExcelDataImportAndLookController : Controller
    {
        private readonly SchoolUserInfoContext _userInfoContext;

        private readonly NpoiExcelOperationService _excelOperationService;

        /// <summary>
        /// 依赖注入到ioc容器中
        /// </summary>
        /// <param name="schoolUserInfoContext"></param>
        /// <param name="excelOperationService"></param>
        public ExcelDataImportAndLookController(SchoolUserInfoContext schoolUserInfoContext, NpoiExcelOperationService excelOperationService)
        {
            _userInfoContext = schoolUserInfoContext;
            _excelOperationService = excelOperationService;
        }

        // GET: ExcelDataImportAndLook
        public ActionResult Index()
        {
            return View();
        }


        /// <summary>
        /// 获取用户信息
        /// </summary>
        /// <param name="page">当前页码</param>
        /// <param name="limit">显示条数</param>
        /// <param name="userName">用户姓名</param>
        /// <returns></returns>
        public JsonResult GetUserInfo(int page = 1, int limit = 15, string userName = "")
        {
            try
            {
                List<UserInfo> listData;
                var totalCount = 0;
                if (!string.IsNullOrWhiteSpace(userName))
                {
                    listData = _userInfoContext.UserInfos.Where(x => x.UserName.Contains(userName)).OrderByDescending(x => x.Id).Skip((page - 1) * limit).Take(limit).ToList();

                    totalCount = _userInfoContext.UserInfos.Count(x => x.UserName.Contains(userName));
                }
                else
                {
                    listData = _userInfoContext.UserInfos.OrderByDescending(x => x.Id).Skip((page - 1) * limit).Take(limit).ToList();

                    totalCount = _userInfoContext.UserInfos.Count();
                }

                return Json(new { code = 0, count = totalCount, data = listData });
            }
            catch (Exception ex)
            {
                return Json(new { code = 1, msg = ex.Message });
            }
        }


        /// <summary>
        /// 数据导入页面
        /// </summary>
        /// <returns></returns>
        public ActionResult ImportPage()
        {
            return View();
        }

        /// <summary>
        /// 数据导入
        /// </summary>
        /// <param name="file">Form表单文件信息</param>
        /// <returns></returns>
        public JsonResult DataImport(IFormFile file)
        {

            var result = _excelOperationService.ExcelDataBatchImport(file, out string resultMsg);

            return Json(result ? new { code = 1, msg = resultMsg } : new { code = 0, msg = resultMsg });

        }


    }
}