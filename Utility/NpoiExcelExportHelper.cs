/**
 * Author:追逐时光
 * Description：Npoi之Excel数据导出帮助类（创建行列，设置字体样式，单元格边框样式，背景样式，行高，背景色等相关常用样式封装）
 * Description：2020年3月25日
 */
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace YY_Utility
{
    public class NpoiExcelExportHelper
    {
        private static NpoiExcelExportHelper _exportHelper;

        public static NpoiExcelExportHelper _
        {
            get => _exportHelper ?? (_exportHelper = new NpoiExcelExportHelper());
            set => _exportHelper = value;
        }

        /// <summary>
        /// TODO:先创建行，然后在创建对应的列
        /// 创建Excel中指定的行
        /// </summary>
        /// <param name="sheet">Excel工作表对象</param>
        /// <param name="rowNum">创建第几行(从0开始)</param>
        /// <param name="rowHeight">行高</param>
        public HSSFRow CreateRow(ISheet sheet, int rowNum, float rowHeight)
        {
            HSSFRow row = (HSSFRow)sheet.CreateRow(rowNum); //创建行
            row.HeightInPoints = rowHeight; //设置列头行高
            return row;
        }

        /// <summary>
        /// 创建行内指定的单元格
        /// </summary>
        /// <param name="row">需要创建单元格的行</param>
        /// <param name="cellStyle">单元格样式</param>
        /// <param name="cellNum">创建第几个单元格(从0开始)</param>
        /// <param name="cellValue">给单元格赋值</param>
        /// <returns></returns>
        public HSSFCell CreateCells(HSSFRow row, HSSFCellStyle cellStyle, int cellNum, string cellValue)
        {
            HSSFCell cell = (HSSFCell)row.CreateCell(cellNum); //创建单元格
            cell.CellStyle = cellStyle; //将样式绑定到单元格
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                //单元格赋值
                cell.SetCellValue(cellValue);
            }

            return cell;
        }


        /// <summary>
        /// 行内单元格常用样式设置
        /// </summary>
        /// <param name="workbook">Excel文件对象</param>
        /// <param name="hAlignment">水平布局方式</param>
        /// <param name="vAlignment">垂直布局方式</param>
        /// <param name="fontHeightInPoints">字体大小</param>
        /// <param name="isAddBorder">是否需要边框</param>
        /// <param name="boldWeight">字体加粗 (None = 0,Normal = 400，Bold = 700</param>
        /// <param name="fileName">字体（仿宋，楷体，宋体，微软雅黑...与Excel主题字体相对应）</param>
        /// <param name="isAddBorderColor">是否增加边框颜色</param>
        /// <param name="isItalic">是否将文字变为斜体</param>
        /// <param name="isLineFeed">是否自动换行</param>
        /// <param name="isAddCellBackground">是否增加单元格背景颜色</param>
        /// <param name="isAddFillPattern">是否填充图案</param>
        /// <param name="fillPattern">填充图案样式(FineDots 细点，SolidForeground立体前景，isAddFillPattern=true时存在)</param>
        /// <param name="cellBackgroundColor">单元格背景样式（当isAddCellBackground=true时存在）</param>
        /// <param name="fontColor">字体颜色</param>
        /// <param name="underlineStyle">下划线样式（无下划线[None],单下划线[Single],双下划线[Double],会计用单下划线[SingleAccounting],会计用双下划线[DoubleAccounting]）</param>
        /// <param name="typeOffset">字体上标下标(普通默认值[None],上标[Sub],下标[Super]),即字体在单元格内的上下偏移量</param>
        /// <param name="isStrikeout">是否显示删除线</param>
        /// <returns></returns>
        public HSSFCellStyle CreateStyle(HSSFWorkbook workbook, HorizontalAlignment hAlignment, VerticalAlignment vAlignment, short fontHeightInPoints, bool isAddBorder, short boldWeight, string fileName = "宋体", bool isAddBorderColor = true, bool isItalic = false, bool isLineFeed = false, bool isAddCellBackground = false, bool isAddFillPattern = false, FillPattern fillPattern = FillPattern.NoFill, short cellBackgroundColor = HSSFColor.Yellow.Index, short fontColor = HSSFColor.Black.Index, FontUnderlineType underlineStyle =
            FontUnderlineType.None, FontSuperScript typeOffset = FontSuperScript.None, bool isStrikeout = false)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头样式
            cellStyle.Alignment = hAlignment; //水平居中
            cellStyle.VerticalAlignment = vAlignment; //垂直居中
            cellStyle.WrapText = isLineFeed;//自动换行


            //背景颜色，边框颜色，字体颜色都是使用 HSSFColor属性中的对应调色板索引，关于 HSSFColor 颜色索引对照表,详情参考：https://blog.csdn.net/FOR_CHINA2012/article/details/29844661

            //背景颜色，NPOI之Excel——设置单元格背景色 ，详情参考：https://www.cnblogs.com/mq0036/p/9835975.html

            //TODO：引用了NPOI后可通过ICellStyle 接口的 FillForegroundColor 属性实现 Excel 单元格的背景色设置，FillPattern 为单元格背景色的填充样式

            if (isAddCellBackground)
            {
                cellStyle.FillForegroundColor = cellBackgroundColor;//背景样式设置
            }


            if (isAddFillPattern) //是否填充图案
            {
                cellStyle.FillPattern = fillPattern;//填充图案样式(FineDots 细点，SolidForeground立体前景)
            }

            //是否增加边框
            if (isAddBorder)
            {
                //常用的边框样式 None(没有),Thin(细边框，瘦的),Medium(中等),Dashed(虚线),Dotted(星罗棋布的),Thick(厚的),Double(双倍),Hair(头发)[上右下左顺序设置]
                cellStyle.BorderBottom = BorderStyle.Thin;
                cellStyle.BorderRight = BorderStyle.Thin;
                cellStyle.BorderTop = BorderStyle.Thin;
                cellStyle.BorderLeft = BorderStyle.Thin;
            }

            //是否设置边框颜色
            if (isAddBorderColor)
            {
                //边框颜色[上右下左顺序设置]
                cellStyle.TopBorderColor = HSSFColor.DarkGreen.Index;//DarkGreen(黑绿色)
                cellStyle.RightBorderColor = HSSFColor.DarkGreen.Index;
                cellStyle.BottomBorderColor = HSSFColor.DarkGreen.Index;
                cellStyle.LeftBorderColor = HSSFColor.DarkGreen.Index;
            }

            /**
             * 设置相关字体样式
             */
            var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建字体

            //假如字体大小只需要是粗体的话直接使用下面该属性即可
            //cellStyleFont.IsBold = true;

            cellStyleFont.Boldweight = boldWeight; //字体加粗
            cellStyleFont.FontHeightInPoints = fontHeightInPoints; //字体大小
            cellStyleFont.FontName = fileName;//字体（仿宋，楷体，宋体 ）
            cellStyleFont.Color = fontColor;//设置字体颜色
            cellStyleFont.IsItalic = isItalic;//是否将文字变为斜体
            cellStyleFont.Underline = underlineStyle;//字体下划线
            cellStyleFont.TypeOffset = typeOffset;//字体上标下标
            cellStyleFont.IsStrikeout = isStrikeout;//是否有删除线

            cellStyle.SetFont(cellStyleFont); //将字体绑定到样式
            return cellStyle;
        }

        //TODO:关于Excel行列单元格合并问题
        #region
        // //一、首先创建工作表，也就是Excel中的sheet，给工作表赋一个名称(Excel底部名称)
        // var sheet = workbook.CreateSheet("Excel底部名称");

        // //二、指定合并的行列
        // /**
        //  第一个参数：从第几行开始合并
        //  第二个参数：到第几行结束合并
        //  第三个参数：从第几列开始合并
        //  第四个参数：到第几列结束合并
        //**/
        // CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
        // sheet.AddMergedRegion(region);

        //  //三、设置单元格的列宽sheet.SetColumnWidth(单元格索引,1000);//设置对应列宽（单元格索引从0开始，后面接宽度）
        #endregion

        #region 使用simpleDemo

        //private readonly IHostingEnvironment _environment;//依赖注入

        ///// <summary>
        ///// 数据导出简单示例
        ///// </summary>
        ///// <param name="dataSource">数据源(list集合数据源)</param>
        ///// <param name="excelName">Excel导出名称</param>
        ///// <returns></returns>
        //public Infrastructure.Response DataExportSimpleDemo(dynamic dataSource,string excelName)
        //{
        //    var result = new Infrastructure.Response();
        //    //Code=400导出失败，Code=200表示成功

        //    var getDataList = dataSource;//导出数据源

        //    try
        //    {
        //        if (getDataList.Any())
        //        {
        //            //首先创建Excel文件对象
        //            var workbook = new HSSFWorkbook();

        //            //创建工作表，也就是Excel中的sheet，给工作表赋一个名称(Excel底部名称)
        //            var sheet = workbook.CreateSheet("xxxxx表");

        //            //sheet.DefaultColumnWidth = 20;//默认列宽

        //            sheet.ForceFormulaRecalculation = true;//TODO:是否开始Excel导出后公式仍然有效（非必须）

        //            #region table 表格内容设置

        //            #region 标题样式

        //            //设置顶部大标题样式
        //            var cellStyleFont = NpoiExcelExportHelper._.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 15, true, 700);

        //            //第一行表单
        //            var row = NpoiExcelExportHelper._.CreateRow(sheet, 0, 20);

        //            var cell = row.CreateCell(0);
        //            //合并单元格 例： 第1行到第2行 第3列到第4列围成的矩形区域

        //            //TODO:关于Excel行列单元格合并问题
        //            /**
        //              第一个参数：从第几行开始合并
        //              第二个参数：到第几行结束合并
        //              第三个参数：从第几列开始合并
        //              第四个参数：到第几列结束合并
        //            **/
        //            CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
        //            sheet.AddMergedRegion(region);

        //            cell.SetCellValue("年度数据汇总表");//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
        //            cell.CellStyle = cellStyleFont;

        //            //二级标题列样式设置
        //            var headTopStyle = NpoiExcelExportHelper._.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 10, true, 700);

        //            //表头名称
        //            var headerName = new[] { "表头一", "表头二", "表头三", "表头四" };

        //            row = NpoiExcelExportHelper._.CreateRow(sheet, 1, 20);//第二行
        //            for (var i = 0; i < headerName.Length; i++)
        //            {
        //                cell = NpoiExcelExportHelper._.CreateCells(row, headTopStyle, i, headerName[i]);

        //                //设置单元格宽度
        //                sheet.SetColumnWidth(i, 10000);//设置对应列宽
        //            }
        //            #endregion


        //            #region 单元格内容信息

        //            //单元格边框样式
        //            var cellStyle = NpoiExcelExportHelper._.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 10, true, 400);

        //            for (var i = 0; i < getDataList.Count; i++)
        //            {
        //                row = NpoiExcelExportHelper._.CreateRow(sheet, i + 2, 20); //sheet.CreateRow(i+2);//在上面表头的基础上创建行

        //                //创建单元格列
        //                cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 0, getDataList[i].AreaName);
        //                cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 1, getDataList[i].MainIndicatorsTourism);
        //                cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 2, getDataList[i].ReportForOvernightTourists);
        //                cell = NpoiExcelExportHelper._.CreateCells(row, cellStyle, 3, getDataList[i].FillingStatusByCountry);

        //            }
        //            #endregion

        //            #endregion

        //            string folder = DateTime.Now.ToString("yyyyMMdd");


        //            //保存文件到静态资源文件夹中（wwwroot）,使用绝对路径
        //            var uploadPath = _environment.WebRootPath+"/Upload/" + folder + "/";

        //            //excel保存文件名
        //            string excelFileName=excelName+"_" + DateTime.Now.ToString("yyyyMMddHHmmss")+".xls";

        //            //创建目录文件夹
        //            if (!Directory.Exists(uploadPath))
        //            {
        //                Directory.CreateDirectory(uploadPath);
        //            }

        //            //Excel的路径及名称
        //            string excelPath = uploadPath + excelFileName;

        //            //使用FileStream文件流来写入数据（传入参数为：文件所在路径，对文件的操作方式，对文件内数据的操作）
        //            var fileStream = new FileStream(excelPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

        //            //向Excel文件对象写入文件流，生成Excel文件
        //            workbook.Write(fileStream);

        //            //关闭文件流
        //            fileStream.Close();

        //            //释放流所占用的资源
        //            fileStream.Dispose();

        //            //excel文件保存的相对路径，提供前端下载
        //            var relativePositioning = "/Upload/" + folder + "/" + excelFileName;

        //            result.Message = relativePositioning;
        //        }
        //        else
        //        {
        //            result.Code = 400;
        //            result.Message = "暂未发现有导出的数据~";
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        result.Code = 400;
        //        result.Message = e.Message;
        //    }
        //    return result;
        //}


        #endregion
    }
}
