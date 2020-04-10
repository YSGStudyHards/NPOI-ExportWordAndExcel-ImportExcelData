using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using YY_Dal;
using YY_Model;
using YY_Utility;

namespace YY_Services
{
    /// <summary>
    /// Excel文档生成并保存,和Excel文档中的数据批量导出操作类
    /// </summary>
    public class NpoiExcelOperationService
    {
        private static IHostingEnvironment _environment;

        private readonly SchoolUserInfoContext _shoSchoolUserInfoContext;

        public NpoiExcelOperationService(SchoolUserInfoContext schoolUserInfoContext, IHostingEnvironment iEnvironment)
        {
            _shoSchoolUserInfoContext = schoolUserInfoContext;
            _environment = iEnvironment;
        }

        /// <summary>
        /// Excel数据导出简单示例
        /// </summary>
        /// <param name="resultMsg">导出结果</param>
        /// <param name="excelFilePath">保存excel文件路径</param>
        /// <returns></returns>
        public bool ExcelDataExport(out string resultMsg, out string excelFilePath)
        {
            var result = true;
            excelFilePath = "";
            resultMsg = "successfully";
            //Excel导出名称
            string excelName = "人才培训课程表";
            try
            {
                //首先创建Excel文件对象
                var workbook = new HSSFWorkbook();

                //创建工作表，也就是Excel中的sheet，给工作表赋一个名称(Excel底部名称)
                var sheet = workbook.CreateSheet("人才培训课程表");

                //sheet.DefaultColumnWidth = 20;//默认列宽

                sheet.ForceFormulaRecalculation = true;//TODO:是否开始Excel导出后公式仍然有效（非必须）

                #region table 表格内容设置

                #region 标题样式

                //设置顶部大标题样式
                var cellStyleFont = NpoiExcelExportHelper._.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 20, true, 700, "楷体", true, false, false, true, FillPattern.SolidForeground, HSSFColor.Coral.Index, HSSFColor.White.Index,
                    FontUnderlineType.None, FontSuperScript.None, false);

                //第一行表单
                var row = NpoiExcelExportHelper._.CreateRow(sheet, 0, 28);

                var cell = row.CreateCell(0);
                //合并单元格 例： 第1行到第2行 第3列到第4列围成的矩形区域

                //TODO:关于Excel行列单元格合并问题
                /**
                  第一个参数：从第几行开始合并
                  第二个参数：到第几行结束合并
                  第三个参数：从第几列开始合并
                  第四个参数：到第几列结束合并
                **/
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 5);
                sheet.AddMergedRegion(region);

                cell.SetCellValue("人才培训课程表");//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
                cell.CellStyle = cellStyleFont;

                //二级标题列样式设置
                var headTopStyle = NpoiExcelExportHelper._.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 15, true, 700, "楷体", true, false, false, true, FillPattern.SolidForeground, HSSFColor.Grey25Percent.Index, HSSFColor.Black.Index,
                FontUnderlineType.None, FontSuperScript.None, false);

                //表头名称
                var headerName = new[] { "课程类型", "序号", "日期", "课程名称", "内容概要", "讲师简介" };

                row = NpoiExcelExportHelper._.CreateRow(sheet, 1, 24);//第二行
                for (var i = 0; i < headerName.Length; i++)
                {
                    cell = NpoiExcelExportHelper._.CreateCells(row, headTopStyle, i, headerName[i]);

                    //设置单元格宽度
                    if (headerName[i] == "讲师简介" || headerName[i] == "内容概要")
                    {
                        sheet.SetColumnWidth(i, 10000);
                    }
                    else

                    {
                        sheet.SetColumnWidth(i, 5000);
                    }

                }
                #endregion


                #region 单元格内容信息

                //单元格边框样式
                var cellStyle = NpoiExcelExportHelper._.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 10, true, 400);

                //左侧列单元格合并 begin
                //TODO:关于Excel行列单元格合并问题（合并单元格后，只需对第一个位置赋值即可）
                /**
                  第一个参数：从第几行开始合并
                  第二个参数：到第几行结束合并
                  第三个参数：从第几列开始合并
                  第四个参数：到第几列结束合并
                **/
                CellRangeAddress leftOne = new CellRangeAddress(2, 7, 0, 0);

                sheet.AddMergedRegion(leftOne);

                CellRangeAddress leftTwo = new CellRangeAddress(8, 11, 0, 0);

                sheet.AddMergedRegion(leftTwo);

                //左侧列单元格合并 end

                var currentDate = DateTime.Now;

                string[] curriculumList = new[] { "艺术学", "设计学", "材料学", "美学", "心理学", "中国近代史", "管理人员的情绪修炼", "高效时间管理", "有效的目标管理", "沟通与协调" };

                int number = 1;

                for (var i = 0; i < 10; i++)
                {
                    row = NpoiExcelExportHelper._.CreateRow(sheet, i + 2, 20); //sheet.CreateRow(i+2);//在上面表头的基础上创建行
                    switch (number)
                    {
                        case 1:
                            cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 0, "公共类课程");
                            break;
                        case 7:
                            cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 0, "管理类课程");
                            break;
                    }

                    //创建单元格列公众类课程
                    cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 1, number.ToString());
                    cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 2, currentDate.AddDays(number).ToString("yyyy-MM-dd"));
                    cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 3, curriculumList[i]);
                    cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 4, "提升，充实，拓展自己综合实力");
                    cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 5, "追逐时光_" + number + "号金牌讲师！");

                    number++;
                }
                #endregion

                #endregion

                string folder = DateTime.Now.ToString("yyyyMMdd");


                //保存文件到静态资源文件夹中（wwwroot）,使用绝对路径
                var uploadPath = _environment.WebRootPath + "/UploadFile/" + folder + "/";

                //excel保存文件名
                string excelFileName = excelName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

                //创建目录文件夹
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                //Excel的路径及名称
                string excelPath = uploadPath + excelFileName;

                //使用FileStream文件流来写入数据（传入参数为：文件所在路径，对文件的操作方式，对文件内数据的操作）
                var fileStream = new FileStream(excelPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                //向Excel文件对象写入文件流，生成Excel文件
                workbook.Write(fileStream);

                //关闭文件流
                fileStream.Close();

                //释放流所占用的资源
                fileStream.Dispose();

                //excel文件保存的相对路径，提供前端下载
                var relativePositioning = "/UploadFile/" + folder + "/" + excelFileName;

                excelFilePath = relativePositioning;
            }
            catch (Exception e)
            {
                result = false;
                resultMsg = e.Message;
            }
            return result;
        }


        /// <summary>
        /// Excel中的数据批量导入
        /// </summary>
        /// <param name="formFile">表单文件信息</param>
        /// <param name="resultMsg">导入结果</param>
        /// <returns></returns>
        public bool ExcelDataBatchImport(IFormFile formFile, out string resultMsg)
        {
            resultMsg = "导入成功";
            var result = false;

            try
            {
                if (formFile.Length > 0)
                {
                    var stopWatch = new Stopwatch();

                    stopWatch.Start();
                    //将excel表格中的数据转化为dataTable数据
                    var getDataTable = NpoiExcelImportHelper._.ExcelToDataTable(formFile.OpenReadStream(), Path.GetExtension(formFile.FileName), out result, out resultMsg
                    );

                    if (getDataTable.Rows.Count <= 0)
                    {
                        return false;
                    }

                    var userInfoList = new List<UserInfo>();

                    var random = DateTime.Now.ToString("fff");

                    //将dataTable数据源转化为List数据源
                    for (int i =0; i < getDataTable.Rows.Count; i++)
                    {
                        var userInfo = new UserInfo
                        {
                            UserName = getDataTable.Rows[i][0].ToString()+"_"+random,
                            Sex = getDataTable.Rows[i][1].ToString(),
                            Phone = getDataTable.Rows[i][2].ToString(),
                            Description = getDataTable.Rows[i][3].ToString(),
                            Hobby = getDataTable.Rows[i][4].ToString()
                        };

                        userInfoList.Add(userInfo);
                    }

                    //EF之AddRange批量添加数据

                    _shoSchoolUserInfoContext.UserInfos.AddRange(userInfoList);

                    _shoSchoolUserInfoContext.SaveChanges();

                    stopWatch.Stop();

                    resultMsg = resultMsg + $",耗费总时长{stopWatch.Elapsed.TotalSeconds}秒，总共导入{getDataTable.Rows.Count}条数据";
                }
                else
                {
                    resultMsg = "Excel表中无数据可导入！";
                }
            }
            catch (Exception e)
            {
                result = false;
                resultMsg = e.Message;
            }

            return result;
        }
    }
}
