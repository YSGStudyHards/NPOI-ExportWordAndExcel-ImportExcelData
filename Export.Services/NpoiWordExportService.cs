using Microsoft.AspNetCore.Hosting;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace Export.Services
{
    public class NpoiWordExportService
    {
        private static IHostingEnvironment _environment;

        public NpoiWordExportService(IHostingEnvironment iEnvironment)
        {
            _environment = iEnvironment;
        }

        #region 生成word

        /// <summary>
        ///  生成word文档,并保存静态资源文件夹（wwwroot)下的SaveWordFile文件夹中
        /// </summary>
        /// <param name="savePath">保存路径</param>
        public bool SaveWordFile(out string savePath)
        {
            savePath = "";
            try
            {
                string currentDate = DateTime.Now.ToString("yyyyMMdd");
                string checkTime = DateTime.Now.ToString("yyyy年MM月dd日");//检查时间
                //保存文件到静态资源wwwroot,使用绝对路径路径
                var uploadPath = _environment.WebRootPath + "/SaveWordFile/" + currentDate + "/";//>>>相当于HttpContext.Current.Server.MapPath("") 

                string workFileName = checkTime + "追逐时光企业员工培训考核统计记录表";
                string fileName = string.Format("{0}.docx", workFileName, System.Text.Encoding.UTF8);

                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }

                //通过使用文件流，创建文件流对象，向文件流中写入内容，并保存为Word文档格式
                using (var stream = new FileStream(Path.Combine(uploadPath, fileName), FileMode.Create, FileAccess.Write))
                {
                    //创建document文档对象对象实例
                    XWPFDocument document = new XWPFDocument();

                    /**
                     *这里我通过设置公共的Word文档中SetParagraph（段落）实例创建和段落样式格式设置，大大减少了代码的冗余，
                     * 避免每使用一个段落而去创建一次段落实例和设置段落的基本样式
                     *(如下，ParagraphInstanceSetting为段落实例创建和样式设置，后面索引表示为当前是第几行段落,索引从0开始)
                     */
                    //文本标题
                    document.SetParagraph(ParagraphInstanceSetting(document, workFileName, true, 19, "宋体", ParagraphAlignment.CENTER), 0);

                    //TODO:这里一行需要显示两个文本
                    document.SetParagraph(ParagraphInstanceSetting(document, $"编号：20190927101120445887", false, 14, "宋体", ParagraphAlignment.CENTER, true, $"    检查时间：{checkTime}"), 1);


                    document.SetParagraph(ParagraphInstanceSetting(document, "登记机关：企业员工监督检查机构", false, 14, "宋体", ParagraphAlignment.LEFT), 2);


                    #region 文档第一个表格对象实例
                    //创建文档中的表格对象实例
                    XWPFTable firstXwpfTable = document.CreateTable(4, 4);//显示的行列数rows:3行,cols:4列
                    firstXwpfTable.Width = 5200;//总宽度
                    firstXwpfTable.SetColumnWidth(0, 1300); /* 设置列宽 */
                    firstXwpfTable.SetColumnWidth(1, 1100); /* 设置列宽 */
                    firstXwpfTable.SetColumnWidth(2, 1400); /* 设置列宽 */
                    firstXwpfTable.SetColumnWidth(3, 1400); /* 设置列宽 */

                    //Table 表格第一行展示...后面的都是一样，只改变GetRow中的行数
                    firstXwpfTable.GetRow(0).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "企业名称", ParagraphAlignment.CENTER, 40, true));
                    firstXwpfTable.GetRow(0).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "追逐时光", ParagraphAlignment.CENTER, 40, false));
                    firstXwpfTable.GetRow(0).GetCell(2).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "企业地址", ParagraphAlignment.CENTER, 40, true));
                    firstXwpfTable.GetRow(0).GetCell(3).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "湖南省-长沙市-岳麓区", ParagraphAlignment.CENTER, 40, false));

                    //Table 表格第二行
                    firstXwpfTable.GetRow(1).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "联系人", ParagraphAlignment.CENTER, 40, true));
                    firstXwpfTable.GetRow(1).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "小明同学", ParagraphAlignment.CENTER, 40, false));
                    firstXwpfTable.GetRow(1).GetCell(2).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "联系方式", ParagraphAlignment.CENTER, 40, true));
                    firstXwpfTable.GetRow(1).GetCell(3).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "151****0456", ParagraphAlignment.CENTER, 40, false));


                    //Table 表格第三行
                    firstXwpfTable.GetRow(2).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "企业许可证号", ParagraphAlignment.CENTER, 40, true));
                    firstXwpfTable.GetRow(2).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "XXXXX-66666666", ParagraphAlignment.CENTER, 40, false));
                    firstXwpfTable.GetRow(2).GetCell(2).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "检查次数", ParagraphAlignment.CENTER, 40, true));
                    firstXwpfTable.GetRow(2).GetCell(3).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, $"本年度检查8次", ParagraphAlignment.CENTER, 40, false));


                    firstXwpfTable.GetRow(3).MergeCells(0, 3);//合并3列
                    firstXwpfTable.GetRow(3).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "", ParagraphAlignment.LEFT, 10, false));

                    #endregion

                    var checkPeopleNum = 0;//检查人数
                    var totalScore = 0;//总得分

                    #region 文档第二个表格对象实例（遍历表格项）
                    //创建文档中的表格对象实例
                    XWPFTable secoedXwpfTable = document.CreateTable(5, 4);//显示的行列数rows:8行,cols:4列
                    secoedXwpfTable.Width = 5200;//总宽度
                    secoedXwpfTable.SetColumnWidth(0, 1300); /* 设置列宽 */
                    secoedXwpfTable.SetColumnWidth(1, 1100); /* 设置列宽 */
                    secoedXwpfTable.SetColumnWidth(2, 1400); /* 设置列宽 */
                    secoedXwpfTable.SetColumnWidth(3, 1400); /* 设置列宽 */

                    //遍历表格标题
                    secoedXwpfTable.GetRow(0).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "员工姓名", ParagraphAlignment.CENTER, 40, true));
                    secoedXwpfTable.GetRow(0).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "性别", ParagraphAlignment.CENTER, 40, true));
                    secoedXwpfTable.GetRow(0).GetCell(2).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "年龄", ParagraphAlignment.CENTER, 40, true));
                    secoedXwpfTable.GetRow(0).GetCell(3).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "综合评分", ParagraphAlignment.CENTER, 40, true));

                    //遍历四条数据
                    for (var i = 1; i < 5; i++)
                    {
                        secoedXwpfTable.GetRow(i).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "小明" + i + "号", ParagraphAlignment.CENTER, 40, false));
                        secoedXwpfTable.GetRow(i).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, "男", ParagraphAlignment.CENTER, 40, false));
                        secoedXwpfTable.GetRow(i).GetCell(2).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, 20 + i + "岁", ParagraphAlignment.CENTER, 40, false));
                        secoedXwpfTable.GetRow(i).GetCell(3).SetParagraph(SetTableParagraphInstanceSetting(document, firstXwpfTable, 90 + i + "分", ParagraphAlignment.CENTER, 40, false));

                        checkPeopleNum++;
                        totalScore += 90 + i;
                    }

                    #endregion

                    #region 文档第三个表格对象实例
                    //创建文档中的表格对象实例
                    XWPFTable thirdXwpfTable = document.CreateTable(5, 4);//显示的行列数rows:5行,cols:4列
                    thirdXwpfTable.Width = 5200;//总宽度
                    thirdXwpfTable.SetColumnWidth(0, 1300); /* 设置列宽 */
                    thirdXwpfTable.SetColumnWidth(1, 1100); /* 设置列宽 */
                    thirdXwpfTable.SetColumnWidth(2, 1400); /* 设置列宽 */
                    thirdXwpfTable.SetColumnWidth(3, 1400); /* 设置列宽 */
                    //Table 表格第一行，后面的合并3列(注意关于表格中行合并问题，先合并，后填充内容)
                    thirdXwpfTable.GetRow(0).MergeCells(0, 3);//从第一列起,合并3列
                    thirdXwpfTable.GetRow(0).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "检查内容: " +
                        $"于{checkTime}下午检查了追逐时光企业员工培训考核并对员工的相关信息进行了相关统计，统计结果如下：                                                                                                                                                                                                                " +
                        "-------------------------------------------------------------------------------------" +
                        $"共对该企业（{checkPeopleNum}）人进行了培训考核，培训考核总得分为（{totalScore}）分。 " + "", ParagraphAlignment.LEFT, 30, false));


                    //Table 表格第二行
                    thirdXwpfTable.GetRow(1).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "检查结果: ", ParagraphAlignment.CENTER, 40, true));
                    thirdXwpfTable.GetRow(1).MergeCells(1, 3);//从第二列起，合并三列
                    thirdXwpfTable.GetRow(1).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "该企业非常优秀，坚持每天学习打卡，具有蓬勃向上的活力。", ParagraphAlignment.LEFT, 40, false));

                    //Table 表格第三行
                    thirdXwpfTable.GetRow(2).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "处理结果: ", ParagraphAlignment.CENTER, 40, true));
                    thirdXwpfTable.GetRow(2).MergeCells(1, 3);
                    thirdXwpfTable.GetRow(2).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "通过检查，评分为优秀！", ParagraphAlignment.LEFT, 40, false));

                    //Table 表格第四行，后面的合并3列(注意关于表格中行合并问题，先合并，后填充内容),额外说明
                    thirdXwpfTable.GetRow(3).MergeCells(0, 3);//合并3列
                    thirdXwpfTable.GetRow(3).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "备注说明: 记住，坚持就是胜利，永远保持一种求知，好问的心理！", ParagraphAlignment.LEFT, 30, false));

                    //Table 表格第五行
                    thirdXwpfTable.GetRow(4).MergeCells(0, 1);
                    thirdXwpfTable.GetRow(4).GetCell(0).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "                                                                                                                                                                                                 检查人员签名：              年 月 日", ParagraphAlignment.LEFT, 40, false));
                    thirdXwpfTable.GetRow(4).MergeCells(1, 2);

                    thirdXwpfTable.GetRow(4).GetCell(1).SetParagraph(SetTableParagraphInstanceSetting(document, thirdXwpfTable, "                                                                                                                                                                                                 企业法人签名：              年 月 日", ParagraphAlignment.LEFT, 40, false));


                    #endregion

                    //向文档流中写入内容，生成word
                    document.Write(stream);

                    savePath = "/SaveWordFile/" + currentDate + "/" + fileName;

                    return true;
                }
            }
            catch (Exception ex)
            {
                //ignore
                savePath = ex.Message;
                return false;
            }
        }


        /// <summary>
        /// 创建word文档中的段落对象和设置段落文本的基本样式（字体大小，字体，字体颜色，字体对齐位置）
        /// </summary>
        /// <param name="document">document文档对象</param>
        /// <param name="fillContent">段落第一个文本对象填充的内容</param>
        /// <param name="isBold">是否加粗</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="fontFamily">字体</param>
        /// <param name="paragraphAlign">段落排列（左对齐，居中，右对齐）</param>
        /// <param name="isStatement">是否在同一段落创建第二个文本对象（解决同一段落里面需要填充两个或者多个文本值的情况，多个文本需要自己拓展，现在最多支持两个）</param>
        /// <param name="secondFillContent">第二次声明的文本对象填充的内容，样式与第一次的一致</param>
        /// <returns></returns>
        private static XWPFParagraph ParagraphInstanceSetting(XWPFDocument document, string fillContent, bool isBold, int fontSize, string fontFamily, ParagraphAlignment paragraphAlign, bool isStatement = false, string secondFillContent = "")
        {
            XWPFParagraph paragraph = document.CreateParagraph();//创建段落对象
            paragraph.Alignment = paragraphAlign;//文字显示位置,段落排列（左对齐，居中，右对齐）

            XWPFRun xwpfRun = paragraph.CreateRun();//创建段落文本对象
            xwpfRun.IsBold = isBold;//文字加粗
            xwpfRun.SetText(fillContent);//填充内容
            xwpfRun.FontSize = fontSize;//设置文字大小
            xwpfRun.SetFontFamily(fontFamily, FontCharRange.None); //设置标题样式如：（微软雅黑，隶书，楷体）根据自己的需求而定

            if (isStatement)
            {
                XWPFRun secondxwpfRun = paragraph.CreateRun();//创建段落文本对象
                secondxwpfRun.IsBold = isBold;//文字加粗
                secondxwpfRun.SetText(secondFillContent);//填充内容
                secondxwpfRun.FontSize = fontSize;//设置文字大小
                secondxwpfRun.SetFontFamily(fontFamily, FontCharRange.None); //设置标题样式如：（微软雅黑，隶书，楷体）根据自己的需求而定
            }


            return paragraph;
        }

        /// <summary>  
        /// 创建Word文档中表格段落实例和设置表格段落文本的基本样式（字体大小，字体，字体颜色，字体对齐位置）
        /// </summary>  
        /// <param name="document">document文档对象</param>  
        /// <param name="table">表格对象</param>  
        /// <param name="fillContent">要填充的文字</param>  
        /// <param name="paragraphAlign">段落排列（左对齐，居中，右对齐）</param>
        /// <param name="rowsHeight">设置文本位置（设置两行之间的行间），从而实现table的高度设置效果  </param>
        /// <param name="isBold">是否加粗（true加粗，false不加粗）</param>
        /// <param name="fontSize">字体大小</param>
        /// <returns></returns>  
        private static XWPFParagraph SetTableParagraphInstanceSetting(XWPFDocument document, XWPFTable table, string fillContent, ParagraphAlignment paragraphAlign, int rowsHeight, bool isBold, int fontSize = 10)
        {
            var para = new CT_P();
            XWPFParagraph paragraph = new XWPFParagraph(para, table.Body);//创建表格中的段落对象
            paragraph.Alignment = paragraphAlign;//文字显示位置,段落排列（左对齐，居中，右对齐）

            XWPFRun xwpfRun = paragraph.CreateRun();//创建段落文本对象
            xwpfRun.SetText(fillContent);
            xwpfRun.FontSize = fontSize;//字体大小
            xwpfRun.IsBold = isBold;//是否加粗
            xwpfRun.SetFontFamily("宋体", FontCharRange.None);//设置字体（如：微软雅黑,华文楷体,宋体）
            xwpfRun.SetTextPosition(rowsHeight);//设置文本位置（设置两行之间的行间），从而实现table的高度设置效果 
            return paragraph;
        }

        #endregion


    }
}
