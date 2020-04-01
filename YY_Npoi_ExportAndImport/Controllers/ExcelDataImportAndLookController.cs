using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.AspNetCore.Http;
using YY_Dal;
using ActionResult = Microsoft.AspNetCore.Mvc.ActionResult;
using Controller = Microsoft.AspNetCore.Mvc.Controller;
using JsonResult = Microsoft.AspNetCore.Mvc.JsonResult;

namespace YY_Npoi_ExportAndImport.Controllers
{
    public class ExcelDataImportAndLookController : Controller
    {
        private readonly SchoolUserInfoContext _userInfoContext;



        /// <summary>
        /// 依赖注入到ioc容器中
        /// </summary>
        /// <param name="schoolUserInfoContext"></param>
        public ExcelDataImportAndLookController(SchoolUserInfoContext schoolUserInfoContext)
        {
            _userInfoContext = schoolUserInfoContext;
        }

        // GET: ExcelDataImportAndLook
        public ActionResult Index()
        {
            return View();
        }


        /// <summary>
        /// 获取用户信息
        /// </summary>
        /// <param name="page"></param>
        /// <param name="limit"></param>
        /// <returns></returns>
        public JsonResult GetUserInfo(int page = 1, int limit = 15)
        {
            try
            {
                //使用ef--skip().take()进行数据分页前面必须增加orderby，否则报错
                var listData = _userInfoContext.UserInfos.OrderBy(p => p.Id).Skip((page - 1) * limit).Take(limit).ToList();

                return Json(new { code = 0, count = _userInfoContext.UserInfos.Count(), data = listData });
            }
            catch (Exception ex)
            {
                return Json(new { code = 1, msg = ex.Message });
            }
        }

        // GET: ExcelDataImportAndLook/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: ExcelDataImportAndLook/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ExcelDataImportAndLook/Create
        [Microsoft.AspNetCore.Mvc.HttpPost]
        [Microsoft.AspNetCore.Mvc.ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: ExcelDataImportAndLook/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: ExcelDataImportAndLook/Edit/5
        [Microsoft.AspNetCore.Mvc.HttpPost]
        [Microsoft.AspNetCore.Mvc.ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: ExcelDataImportAndLook/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: ExcelDataImportAndLook/Delete/5
        [Microsoft.AspNetCore.Mvc.HttpPost]
        [Microsoft.AspNetCore.Mvc.ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }
    }
}