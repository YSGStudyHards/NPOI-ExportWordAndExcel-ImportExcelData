/**
 * Author:追逐时光
 * Description：Npoi数据导入帮助类
 * Description：2020年4月5日
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace YY_Utility
{
    public class NpoiExcelImportHelper
    {

        private static NpoiExcelImportHelper _excelImportHelper;

        public static NpoiExcelImportHelper _
        {
            get => _excelImportHelper ?? (_excelImportHelper = new NpoiExcelImportHelper());
            set => _excelImportHelper = value;
        }

        /// <summary>
        /// 读取excel表格中的数据,将Excel文件流转化为dataTable数据源  
        /// 默认第一行为标题 
        /// </summary>
        /// <param name="stream">excel文档文件流</param>
        /// <param name="fileType">文档格式</param>
        /// <param name="isSuccess">是否转化成功</param>
        /// <param name="resultMsg">转换结果消息</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(Stream stream, string fileType, out bool isSuccess, out string resultMsg)
        {
            isSuccess = false;
            resultMsg = "Excel文件流成功转化为DataTable数据源";
            var excelToDataTable = new DataTable();

            try
            {
                IWorkbook workbook;

                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                #region 判断excel版本
                switch (fileType)
                {
                    //2007以下版本excel
                    case ".xlsx":
                        workbook = new XSSFWorkbook(stream);
                        break;
                    case ".xls":
                        workbook = new HSSFWorkbook(stream);
                        break;
                    default:
                        throw new Exception("Excel文档格式有误");
                }
                #endregion

                var sheet = workbook.GetSheetAt(0);
                var rows = sheet.GetRowEnumerator();

                var headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;//最后一行列数（即为总列数）

                //获取第一行标题列数据源,转换为dataTable数据源的表格标题名称
                for (var j = 0; j < cellCount; j++)
                {
                    var cell = headerRow.GetCell(j);
                    excelToDataTable.Columns.Add(cell.ToString());
                }

                //获取Excel表格中除标题以为的所有数据源，转化为dataTable中的表格数据源
                for (var i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    var dataRow = excelToDataTable.NewRow();

                    var row = sheet.GetRow(i);

                    if (row == null) continue; //没有数据的行默认是null　

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = row.GetCell(j).ToString();
                    }
                    excelToDataTable.Rows.Add(dataRow);
                }

                isSuccess = true;
            }
            catch (Exception e)
            {
                resultMsg = e.Message;
            }

            return excelToDataTable;
        }
    }
}
