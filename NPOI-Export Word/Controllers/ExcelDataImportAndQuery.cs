using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using YY_Dal;
using YY_Services;

namespace NPOI_Export_Word.Controllers
{
    /// <summary>
    /// Excel数据导入和查询
    /// </summary>
    public class ExcelDataImportAndQuery
    {
        private readonly SchoolUserInfoContext _userInfoContext;

        /// <summary>
        /// 依赖注入到ioc容器中
        /// </summary>
        /// <param name="userInfoContext"></param>
        public ExcelDataImportAndQuery(SchoolUserInfoContext userInfoContext)
        {
            _userInfoContext = userInfoContext;
        }


        /// <summary>
        /// 读取Excel文件中的数据
        /// </summary>  
        /// <param name="strFileName">excel文档路径</param>
        /// <param name="fileType">文件类型</param>
        /// <returns></returns>  
        public static DataTable ExcelDataImport(string strFileName, string fileType)
        {
            IWorkbook workbook;
            DataTable dt = new DataTable();

            //HSSFWorkbook hssfworkbook;
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))//数据读取
            {
                ////XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                //不同格式excle判断
                if (fileType == "xls")
                {
                    workbook = new HSSFWorkbook(file);
                }
                else
                {
                    workbook = new XSSFWorkbook(file);
                }
            }
            ISheet sheet = workbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                dt.Columns.Add(cell.ToString());
            }

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = dt.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                        dataRow[j] = row.GetCell(j).ToString();
                }

                dt.Rows.Add(dataRow);
            }
            return dt;
        }


    }
}
