using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Collections;
using System.Security.AccessControl;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace nsLims_NPOI
{
    public partial class classLims_NPOI : Form
    {
        public classLims_NPOI()
        {
            InitializeComponent();
        }

        //A4纸的固定长度为677.75,倍数为20
        private static double PAGE_HEIGHT = 57 * 14.25 * 20;//-2170;//16245
        //A4纸的像素宽度为84.5,,倍数为278
        private static double PAGE_WIDTH = 84.5;
        //workbook
        private IWorkbook iWorkbook;
        public IWorkbook IWorkbook
        {
            get
            {
                return iWorkbook;
            }

            set
            {
                iWorkbook = value;
            }
        }

        /// <summary>
        /// 将异常打印到LOG文件
        /// </summary>
        /// <param name="ex">异常</param>
        /// <param name="LogAddress">日志文件地址,如果为空,则默认为C:\\Lims_NPOI</param>
        public static void WriteLog(Exception ex, string LogAddress = "")
        {
            //如果日志文件为空，则默认在C:\\Lims_NPOI目录下新建 YYYY-mm-dd_Log.log文件
            if (LogAddress == "")
            {
                LogAddress = "C:\\Lims_NPOI" + '\\' +
                    DateTime.Now.Year + '-' +
                    DateTime.Now.Month + '-' +
                    DateTime.Now.Day + "_Log.log";
            }

            //把异常信息输出到文件
            StreamWriter sw = new StreamWriter(LogAddress, true);
            sw.WriteLine("当前时间：" + DateTime.Now.ToString());
            sw.WriteLine("异常信息：" + ex.Message);
            sw.WriteLine("异常对象：" + ex.Source);
            sw.WriteLine("调用堆栈：\n" + ex.StackTrace.Trim());
            sw.WriteLine("触发方法：" + ex.TargetSite);
            sw.WriteLine();
            sw.Close();
        }

        /// <summary>
        /// 打印文本到log文件
        /// </summary>
        /// <param name="ex">要写入的文本</param>
        /// <param name="LogAddress">Log文件地址,为空时默认在C:\\Lims_NPOI</param>
        public static void WriteLog(string ex, string LogAddress = "")
        {
            //如果日志文件为空，则默认在C:\\Lims_NPOI目录下新建 YYYY-mm-dd_Log.log文件
            if (LogAddress == "")
            {
                LogAddress = "C:\\Lims_NPOI" + '\\' +
                    DateTime.Now.Year + '-' +
                    DateTime.Now.Month + '-' +
                    DateTime.Now.Day + "_Log.log";
            }

            //把异常信息输出到文件
            StreamWriter sw = new StreamWriter(LogAddress, true);
            sw.WriteLine("当前时间：" + DateTime.Now.ToString());
            sw.WriteLine("打印信息：" + ex);
            sw.WriteLine();
            sw.Close();
        }

        #region 禁止使用"C12"作为单元格坐标
        ////返回字符串部分
        ///// <summary>
        ///// 返回字符串部分
        ///// </summary>
        ///// <param name="s">Excel坐标,如B23</param>
        ///// <returns>坐标的字符串部分,实际为列号,如B23返回B</returns>
        //public static string disassemblyToString(string s)
        //{
        //    if (string.IsNullOrEmpty(s)) return "";
        //    int n = 0;
        //    for (int i = s.Length - 1; i >= 0; i--)
        //    {
        //        char c = Char.ToUpper(s[i]);
        //        if (c < 'A' || c > 'Z') n = i;
        //    }
        //    return s.Substring(0, n);
        //}

        ////返回数值部分
        ///// <summary>
        ///// 返回数值部分
        ///// </summary>
        ///// <param name="s">Excel坐标,如B23</param>
        ///// <returns>坐标的字符串部分,实际为列号,如B23返回23</returns>
        //public static int disassemblyToNumber(string s)
        //{
        //    if (string.IsNullOrEmpty(s)) return 0;
        //    int n = 0;
        //    for (int i = 0; i < s.Length; i++)
        //    {
        //        char c = Char.ToUpper(s[i]);
        //        if (c >= '0' && c <= '9')
        //        {
        //            n = i;
        //            break;
        //        }
        //    }
        //    return int.Parse(s.Substring(n));
        //}
        #endregion

        /// <summary>
        /// 设置单元格样式的背景色项
        /// </summary>
        /// <param name="style">源格式</param>
        /// <param name="colorName">颜色名</param>
        /// <returns>目标格式</returns>
        public HSSFCellStyle setCellBGColor(HSSFCellStyle style, string colorName)
        {
            short color = System.Convert.ToInt16(string2ColorIndex(colorName));
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = color;
            return style;
        }

        /// <summary>
        /// 设置单元格字体样式的颜色项
        /// </summary>
        /// <param name="font">源字体</param>
        /// <param name="colorName">颜色名</param>
        /// <returns>目标字体</returns>
        public HSSFFont setCellFontColor(HSSFFont font, string colorName)
        {
            short color = System.Convert.ToInt16(string2ColorIndex(colorName));
            font.Color = color;
            return font;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="workbook">源工作簿</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="col">列号</param>
        /// <param name="width">宽度</param>
        /// <returns>目标工作簿</returns>
        public HSSFWorkbook setColWidth(HSSFWorkbook workbook, int sheetIndex, int col, int width)
        {
            string sheetName = workbook.GetSheetName(sheetIndex);
            return setColWidth(workbook, sheetName, col, width);
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="hssfworkbook">源工作簿</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="col">列号</param>
        /// <param name="width">宽度</param>
        /// <returns>目标工作簿</returns>
        public HSSFWorkbook setColWidth(HSSFWorkbook workbook, string sheetName, int col, int width)
        {
            try
            {
                HSSFSheet sheet = (HSSFSheet)workbook.GetSheet(sheetName);
                sheet.SetColumnWidth(col, width * 288);
                //HSSFRow row = (HSSFRow)sheet.GetRow(0);
                //HSSFCell cell = (HSSFCell)row.GetCell(col);
                //HSSFCellStyle hcs = (HSSFCellStyle)cell.CellStyle;
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }
        }

        /// <summary>
        /// 设置表格列宽,注意由于不知名原因,模板有时不能完整打印,需要重新制作
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sPercent">各列所占表格的百分比</param>
        /// <returns>new sheet</returns>
        public HSSFSheet SetTableColWidth(HSSFSheet sheet, string sPercent)
        {
            try
            {
                string[] aPercent = sPercent.Split(',');
                int firstCol = getColumnRange(sheet)[0];
                int lastCol = getColumnRange(sheet)[1];
                //Excel列数可能会大于数组,此时只设置前部分列宽
                if (lastCol - firstCol + 1 < aPercent.Length)
                {
                    return sheet;
                }
                double sum_aPercent = 0;
                for (int i = 0; i < aPercent.Length; i++)
                {
                    sum_aPercent += Convert.ToDouble(aPercent[i]);
                }
                float[] colWidths = new float[aPercent.Length];
                for (int i = 0; i < aPercent.Length; i++)
                {
                    //A4纸的像素宽度为84.5
                    colWidths[i] = (float)(Convert.ToDouble(aPercent[i]) * PAGE_WIDTH / sum_aPercent);
                }
                for (int i = firstCol; i <= lastCol; i++)
                {
                    //四舍五入取整
                    sheet.SetColumnWidth(i, Convert.ToInt32(colWidths[i - firstCol]) * 288);
                }

                return sheet;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return sheet;
            }
        }

        /// <summary>
        /// 设置打印标题区间
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startRowFlag">起始行所在标记,注意此行不计入表头</param>
        /// <param name="endRowflag">结束行标记</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <returns>2个元素的数组,表头行区间</returns>
        public int[] SetTableHeader(string filePath, int sheetIndex, string startRowFlag, string endRowflag, int startCol, int endCol)
        {
            IWorkbook wb = loadExcelWorkbookI(filePath);
            ISheet sheet = wb.GetSheetAt(sheetIndex);

            //数据写完后会多一个空行,需要手动删掉
            sheet.RemoveRow(sheet.GetRow(sheet.LastRowNum));

            int startRow = selectPosition(sheet, startRowFlag).X + 1;
            int endRow = selectPosition(sheet, endRowflag).X;
            //设置打印标题用,CellRangeAddress参数:(起始行号，终止行号， 起始列号，终止列号)
            sheet.RepeatingRows = new NPOI.SS.Util.CellRangeAddress(startRow, endRow, startCol, endCol);
            saveExcelWithoutAsk(filePath, wb);
            int[] range = new int[] { startRow, endRow };
            return range;
        }

        //返回sheet总页数
        /// <summary>
        /// 返回sheet总页数
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startHead">表头起始行号</param>
        /// <param name="endHead">表头结束行号</param>
        /// <returns>sheet总页数</returns>
        public int getSheetPageCount(string filePath, int sheetIndex, int startHead, int endHead)
        {
            try
            {
                FileConvertClass.excelRefresh(filePath);
                ISheet sheet = loadExcelSheetI(filePath, sheetIndex);
                int fRow = endHead+1;
                int lRow = sheet.LastRowNum;
                double totalH = 0; //总行高
                int pageCount = 0; //页码计数
                double headH = 0; //表头高度
                headH = (sheet.GetMargin(MarginType.TopMargin) + sheet.GetMargin(MarginType.BottomMargin)) / 0.39370078740157483;
                headH = headH * 27.682 * 20;
                for (int i = startHead; i <= endHead; i++)
                {
                    headH += sheet.GetRow(i).Height;
                }

                for (int i = fRow; i <= lRow; i++)
                {
                    IRow row = sheet.GetRow(i);
                    double tempH = row.Height;
                    if (row.ZeroHeight == true)
                    {
                        tempH = 0;
                    }
                    if (totalH + tempH > PAGE_HEIGHT - headH)//超过一页
                    {
                        if (row.Cells.Count <= 0)//当前行为空则结束循环
                        {
                            break;
                        }
                        pageCount++;
                        totalH = tempH;
                    }
                    else
                    {
                        totalH += tempH;
                    }

                }
                return pageCount + 1;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return 0;
            }
        }

       
        /// <summary>
        /// 获取表格总页数,指定sheet要计算表头
        /// </summary>
        /// <param name="filePath">目标表格路径</param>
        /// <param name="sheetIndex">要计算表头的单sheet索引</param>
        /// <param name="startHead">要计算表头的单sheet起始行号</param>
        /// <param name="endHead">要计算表头的单sheet结束行号</param>
        /// <returns>总页数</returns>
        public int getExcelPageCount(string filePath, int sheetIndex, int startHead, int endHead)
        {
            try
            {
                int pages = 0;
                IWorkbook wb = loadExcelWorkbookI(filePath);
                for (int i = 0; i < wb.NumberOfSheets; i++)
                {
                    if (wb.GetSheetAt(i).LastRowNum < 1)
                    {
                        continue;
                    }
                    if (i == sheetIndex)
                    {
                        pages += getSheetPageCount(filePath, sheetIndex, startHead, endHead);
                    }
                    else
                    {
                        //传入起始表头行号和小于它的结束表头行号,代表无表头
                        pages += getSheetPageCount(filePath, sheetIndex, 2, 1);
                    }
                }
                return pages;
            }catch(Exception ex)
            {
                WriteLog(ex, "");
                return 0;
            }
        }

        //设置页码格式
        public void setExcelPageFormat(string filePath, int excPage)
        {
            if (excPage < 0)
            {
                excPage = 0;
            }
            IWorkbook wb = loadExcelWorkbookI(filePath);
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                //空sheet的LastRowNum为0
                if (sheet.LastRowNum < 1)
                {
                    continue;
                }
                string right = "&10共 &N+1 页 第 &P+1 页";
                sheet.Header.Right = right;
                //sheet.Header.Right = "\n共 \"&\"[总页数]+" + (1 + excPage).ToString() + " 页 第 \"&\"[页码]+1 页";
                sheet.Header.Right = "\n&10共 &N+" + (1 + excPage).ToString() + " 页 第 &P+1 页";
            }
            saveExcelWithoutAsk(filePath, wb);
        }

        //设置多个sheet的起始页码
        public IWorkbook setStartPage(IWorkbook wb, int startSheetIndex, string startPage, int totalPages)
        {
            for (int i = startSheetIndex; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                //空sheet的LastRowNum为0
                if (sheet.LastRowNum < 1)
                {
                    continue;
                }
                //sheet.PrintSetup.PageStart = (short)startPageNum;
                sheet.Header.Right = "\n&10共 " + totalPages.ToString() + " 页 第 &P"+ startPage + " 页";
            }
            return wb;
        }

        
        /// <summary>
        /// 拉伸表格到A4大小
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="sheetIndex"></param>
        /// <param name="startHeadRow">表头起始行号</param>
        /// <param name="endHeadRow">表头结束行号</param>
        /// <param name="startCol">起始列号</param>
        /// <param name="endCol">结束列号</param>
        /// <returns>是否成功</returns>
        public bool stretchLastRowHeight(string filePath, int sheetIndex, int startHeadRow, int endHeadRow, int startCol, int endCol)
        {
            try
            {
                //NPOI自动换行后行高读取值不变,需要用COM组件重新保存
                FileConvertClass.excelRefresh(filePath);
                IWorkbook wb = loadExcelWorkbookI(filePath);
                ISheet sheet = wb.GetSheetAt(sheetIndex);
                int pc = stretchLastRowHeight(sheet, startHeadRow, endHeadRow, startCol, endCol);
                saveExcelWithoutAsk(filePath, wb);
                return true;
            }
            catch(Exception ex) 
            {
                WriteLog(ex, "");
                return false;
            }
        }

       
        /// <summary>
        /// 拉伸最后一行,使占满A4纸,并返回总页数
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startHeadRow">表头起始行号</param>
        /// <param name="endHeadRow">表头结束行号</param>
        /// <param name="startCol">起始列号</param>
        /// <param name="endCol">结束列号</param>
        /// <returns>总页数</returns>
        public int stretchLastRowHeight(ISheet sheet, int startHeadRow, int endHeadRow, int startCol, int endCol)
        {
            try
            {
                int fRow = endHeadRow+1;
                int lRow = sheet.LastRowNum;
                
                double totalH = 0; //总行高
                int pageCount = 1; //页码计数
                //高度换算 1厘米＝27.682个单位, 1个单位在NPOI中是20个单位
                double headH = 0; //表头高度                
                headH = (sheet.GetMargin(MarginType.TopMargin) + sheet.GetMargin(MarginType.BottomMargin))/ 0.39370078740157483;
                headH = headH * 27.682*20;
                for (int i = startHeadRow; i <= endHeadRow; i++)
                {
                    if (sheet.GetRow(i).ZeroHeight == true)
                    {
                        headH += 0;
                    }
                    else
                    {
                        headH += sheet.GetRow(i).Height;
                    }
                }

                for (int i = fRow; i <= lRow; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    double tempH;
                    if (row.ZeroHeight == true)
                    {
                        tempH = 0;
                    }
                    else
                    {                        
                        tempH = row.Height;
                    }
                     
                    if (totalH + tempH > PAGE_HEIGHT - (pageCount*headH))//超过一页
                    {
                        pageCount++;
                        if (i == lRow)//已经最后一行
                        {
                            short shH = (short)(PAGE_HEIGHT - totalH - (1* headH));
                            if (shH > 409 * 20)//最高只能设置一行为409
                            {
                                //sheet.ShiftRows(i + 1,                                 //--开始行
                                //    i + 1,                            //--结束行
                                //    1,                             //--移动大小(行数)--往下移动
                                //    true,                                   //是否复制行高
                                //    false,                                  //是否重置行高
                                //    true                                    //是否移动批注
                                //    );
                                IRow newRow = sheet.CreateRow(i + 1);//先新增一行
                                ICell sourceCell = null;
                                ICell targetCell = null;
                                //复制格式到新的行
                                for (int m = row.FirstCellNum; m < row.LastCellNum; m++)
                                {
                                    sourceCell = row.GetCell(m);
                                    if (sourceCell == null)
                                        continue;
                                    targetCell = newRow.CreateCell(m);
                                    targetCell.CellStyle = sourceCell.CellStyle;
                                    targetCell.SetCellType(sourceCell.CellType);

                                }
                                row.Height = 409 * 20;
                                //测试中应该是106*20
                                sheet.GetRow(i + 1).Height = (short)(shH - 409 * 20);
                                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i + 1, startCol, endCol));
                            }
                            else
                            {
                                row.Height = shH;
                            }
                            return pageCount;
                        }
                        else
                        {
                            totalH = tempH;
                            continue;
                        }
                    }
                    else//不超过一页
                    {
                        if (i == lRow)//已经最后一行
                        {
                            short shH = (short)(PAGE_HEIGHT - totalH - (1* headH));
                            if (shH > 409 * 20)//最高只能设置一行为409
                            {
                                //sheet.ShiftRows(i + 1,                                 //--开始行
                                //    i + 1,                            //--结束行
                                //    1,                             //--移动大小(行数)--往下移动
                                //    true,                                   //是否复制行高
                                //    false,                                  //是否重置行高
                                //    true                                    //是否移动批注
                                //    );
                                IRow newRow = sheet.CreateRow(i + 1);//先新增一行
                                ICell sourceCell = null;
                                ICell targetCell = null;
                                //复制格式到新的行
                                for (int m = row.FirstCellNum; m < row.LastCellNum; m++)
                                {
                                    sourceCell = row.GetCell(m);
                                    if (sourceCell == null)
                                        continue;
                                    targetCell = newRow.CreateCell(m);
                                    targetCell.CellStyle = sourceCell.CellStyle;
                                    targetCell.SetCellType(sourceCell.CellType);

                                }
                                row.Height = 409 * 20;
                                //测试中应该是106*20
                                sheet.GetRow(i + 1).Height = (short)(shH - 409 * 20);
                                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i + 1, startCol, endCol));
                            }
                            else
                            {
                                row.Height = shH;
                            }
                            return pageCount;
                        }
                        else
                        {
                            totalH = totalH + tempH;
                            continue;
                        }
                    }
                }
                return pageCount;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return 0;
            }
        }

        
        /// <summary>
        /// 遍历求标记字符串个数
        /// </summary>
        /// <param name="sheet">目标sheet</param>
        /// <param name="flag">目标标记</param>
        /// <returns>标记字符串个数</returns>
        public int getFlagCount(ISheet sheet, string flag)
        {
            int count = 0;
            Point p = new Point();
            p.X = -1;
            p.Y = -1;
            try
            {
                int minRow = sheet.FirstRowNum;
                int maxRow = sheet.LastRowNum;
                for (int i = minRow; i <= maxRow; i++)
                {
                    IRow row = (IRow)sheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    for (int j = row.FirstCellNum; j <= row.LastCellNum; j++)
                    {
                        ICell cell = (ICell)row.GetCell(j);
                        if (cell == null)
                            continue;
                        string cellValue = getCellStringValueAllCase(cell);
                        if (cellValue.IndexOf(flag) > -1)
                        {
                            count++;
                        }
                    }
                }
                return count;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return count;
            }

        }


        /// <summary>
        /// 写入多个图片到模板
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="reportno">报告编号</param>
        /// <param name="ordno">子样号数组</param>
        /// <param name="testcode">检测项数组</param>
        /// <param name="imgDscp">图片描述</param>
        /// <param name="imgPath">图片路径数组</param>
        /// <param name="endFlag">结束标记</param>
        /// <returns>新工作簿</returns>
        public IWorkbook reportImagesExcel(IWorkbook wb, int sheetIndex, 
            string reportno, object[] ordno, object[] testcode, object[] imgDscp, object[] imgPath, string endFlag)
        {
            try
            {
                string[] reportnoList = new string[] { reportno };//报告编号
                string[] ordnoList = dArray2String1(ordno);//子样号
                string[] testcodeList = dArray2String1(testcode);//检测项
                string[] imgDscpList = dArray2String1(imgDscp);//图片描述
                string[] imgPathList = dArray2String1(imgPath);//图片路径
                if(wb.NumberOfSheets<sheetIndex+1)
                {
                    return wb;
                }
                ISheet fromSheet = wb.GetSheetAt(sheetIndex);
                int tempCount = getFlagCount(fromSheet, "&[子样");//查询是1图片模板还是2图片模板
                if (tempCount != 1 && tempCount != 2)
                {
                    return wb;
                }
                //将数据页的结束标记置为空
                else {
                    wb = replaceCellValue(wb, sheetIndex-1, "", endFlag);
                }

                //如果使用2图片模板,奇数图片数组最后一个追加一个""空
                if (tempCount == 2 && imgPathList.Length % 2 == 1)
                {
                    List<string> list = new List<string>(imgPathList);
                    list.Insert(list.Count, "");
                    imgPathList = list.ToArray();
                }

                //添加多个sheet
                for (int i = sheetIndex + 1; i <= (imgPathList.Length - tempCount) / tempCount + sheetIndex; i++)
                {
                    addNewSheet(wb, i, "Sheet" + (i + 1).ToString());
                    CopySheet(wb, fromSheet, wb.GetSheetAt(i), true);                   
                }

                //清除结束标记所在单元格的上边框
                for (int i = sheetIndex; i < (imgPathList.Length - tempCount) / tempCount + sheetIndex; i++)
                {
                    wb = setTopLineNull(wb, i, endFlag);
                    
                }

                //清理结束标记所在单元格的值
                for (int i = sheetIndex; i < (imgPathList.Length - tempCount) / tempCount + sheetIndex; i++)
                {
                    wb = replaceCellValue(wb, i, "", endFlag);
                }

                //填充数据
                for (int i = 0; i < (imgPathList.Length) / tempCount; i++)
                {
                    //如果是1图片模板
                    if (tempCount == 1)
                    {
                        object[] dReportno = { "&[任务编号]", reportnoList[0] };
                        object[] dOrdno = { "&[子样1]", ordnoList[i] };
                        object[] dTestcode = { "&[检测项1]", testcodeList[i] };
                        object[] dImgDscp = { "&[图片说明1]", imgDscpList[i] };
                        object[] dArray = { dReportno, dOrdno, dTestcode, dImgDscp };
                        addImgTo1ImgWorkbook(wb, i + sheetIndex, imgPathList[i], "&[图片1]", dArray);
                    }
                    //如果是2图片模板
                    else if (tempCount == 2)
                    {
                        //如果是奇数数组的最后一页,则静态数据部分只写第一张图片,第二张图片的路径为""
                        if (ordnoList.Length % 2 == 1 && ordnoList.Length <= i*2+1 )
                        {
                            object[] dReportno = { "&[任务编号]", reportnoList[0] };
                            object[] dOrdno1 = { "&[子样1]", ordnoList[i * 2] };
                            object[] dTestcode1 = { "&[检测项1]", testcodeList[i * 2] };
                            object[] dImgDscp1 = { "&[图片说明1]", imgDscpList[i * 2] };
                            object[] dArray = { dReportno, dOrdno1, dTestcode1, dImgDscp1 };
                            addImgTo2ImgWorkbook(wb, i + sheetIndex, imgPathList[i * 2], imgPathList[i * 2 + 1],
                                "&[图片1]", "&[图片2]", endFlag, dArray);
                        }
                        else 
                        {
                            object[] dReportno = { "&[任务编号]", reportnoList[0] };
                            object[] dOrdno1 = { "&[子样1]", ordnoList[i * 2] };
                            object[] dTestcode1 = { "&[检测项1]", testcodeList[i * 2] };
                            object[] dImgDscp1 = { "&[图片说明1]", imgDscpList[i * 2] };
                            object[] dOrdno2 = { "&[子样2]", ordnoList[i * 2 + 1] };
                            object[] dTestcode2 = { "&[检测项2]", testcodeList[i * 2 + 1] };
                            object[] dImgDscp2 = { "&[图片说明2]", imgDscpList[i * 2 + 1] };
                            object[] dArray = { dReportno, dOrdno1, dTestcode1, dImgDscp1, dOrdno2, dTestcode2, dImgDscp2 };
                            addImgTo2ImgWorkbook(wb, i + sheetIndex, imgPathList[i * 2], imgPathList[i * 2 + 1],
                                "&[图片1]", "&[图片2]", endFlag, dArray);
                        }

                    }
                }

                return wb;
            }catch(Exception ex)
            {
                WriteLog(ex, "");
                return wb;
            }

        }

        
        /// <summary>
        /// 写入多个图片到模板
        /// </summary>
        /// <param name="fromPath">源表路径</param>
        /// <param name="toPath">保存路径</param>
        /// <param name="sheetIndex">图片模板所在sheet索引</param>
        /// <param name="reportno">报告编号</param>
        /// <param name="ordno">子样号数组</param>
        /// <param name="testcode">检测项数组</param>
        /// <param name="imgDscp">图片描述</param>
        /// <param name="imgPath">图片路径数组</param>
        /// <param name="endFlag">结束标记</param>
        public void reportImagesExcel(string fromPath, string toPath, int sheetIndex,
            string reportno, object[] ordno, object[] testcode, object[] imgDscp, object[] imgPath, string endFlag)
        {
            try
            {
                IWorkbook wb = loadExcelWorkbookI(fromPath);
                wb = reportImagesExcel(wb, sheetIndex, reportno, ordno, testcode, imgDscp, imgPath, endFlag);
                saveExcelWithoutAsk(toPath, wb);

                return;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return;
            }

        }


        /// <summary>
        /// 设置结束标记所在单元格顶部线为空
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="endFlag">结束标记</param>
        /// <returns>workbook</returns>
        private IWorkbook setTopLineNull(IWorkbook workbook, int sheetIndex, string endFlag)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            int[,] region = getCellMergeArea(sheet, endFlag);
            //CellRangeAddress cra = new CellRangeAddress(region[0, 0], region[0, 1], region[1, 0], region[1, 1]);
            #region 创建合并后单元格风格,黑边框,顶部无边框,水平居中,垂直居中
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.None;//顶部边框不打印
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.LeftBorderColor = HSSFColor.Black.Index;
            style.RightBorderColor = HSSFColor.Black.Index;
            style.TopBorderColor = HSSFColor.White.Index;//顶部边框为默认
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//垂直居中
            style.SetFont(CellUtil.GetRow(region[1,0], sheet).GetCell(region[1,1]).CellStyle.GetFont(workbook));
            #endregion

            for (int i = region[0,0]; i <= region[1,0]; i++)
            {
                IRow row = CellUtil.GetRow(i, sheet);
                for (int j = region[0,1]; j <= region[1,1]; j++)
                {
                    ICell singleCell = HSSFCellUtil.GetCell(row, (short)j);
                    singleCell.CellStyle = style;
                }
            }
            return workbook;
        }

        /// <summary>
        /// 根据NPOI的颜色名,返回其索引值
        /// </summary>
        /// <param name="colorName">字符串形式的颜色名,参考:http://www.holdcode.com/web/details/117 </param>
        /// <returns>NPOI中所对应的颜色索引</returns>
        public int string2ColorIndex(string colorName)
        {
            switch (colorName)
            {
                case "Black": return 8;
                case "White": return 9;
                case "Red": return 10;
                case "BRIGHT_GREEN": return 11;
                case "Blue": return 12;
                case "Yellow": return 13;
                case "Pink": return 14;
                case "TURQUOISE": return 15;
                case "Dark_Red": return 16;
                case "Green": return 17;
                case "Dark_Blue": return 18;
                case "DARK_YELLOW": return 19;
                case "VIOLET": return 20;
                case "Teal": return 21;
                case "GREY_25_PERCENT": return 22;
                case "Grey_50_PERCENT": return 23;
                case "CORNFLOWER_BLUE": return 24;
                case "MAROON": return 25;
                case "LEMON_CHIFFON": return 26;
                case "ORCHID": return 28;
                case "CORAL": return 29;
                case "ROYAL_BLUE": return 30;
                case "LIGHT_CORNFLOWER_BLUE": return 31;
                case "SKY_BLUE": return 40;
                case "LIGHT_TURQUOISE": return 41;
                case "LIGHT_GREEN": return 42;
                case "LIGHT_YELLOW": return 43;
                case "PALE_BLUE": return 44;
                case "Rose": return 45;
                case "LAVENDER": return 46;
                case "Tan": return 47;
                case "LIGHT_BLUE": return 48;
                case "AQUA": return 49;
                case "LIME": return 50;
                case "Gold": return 51;
                case "LIGHT_ORANGE": return 52;
                case "Orange": return 53;
                case "Blue_Grey": return 54;
                case "GREY_40_PERCENT": return 55;
                case "Dark_Teal": return 56;
                case "SEA_GREEN": return 57;
                case "Dark_Green": return 58;
                case "Olive_Green": return 59;
                case "Brown": return 60;
                case "Plum": return 61;
                case "Indigo": return 62;
                case "Grey_80_PERCENT": return 63;
                case "AUTOMATIC": return 64;
                default: return 8;

            }
        }

        ///// <summary>
        ///// 添加一组图片到excel,各个图片间隔1行
        ///// </summary>
        ///// <param name="workbook">源工作簿</param>
        ///// <param name="sheetIndex">工作表sheet索引</param>
        ///// <param name="imagePathList">图片文件路径清单</param>
        ///// <param name="imgCellStr">图片位置标志字符串</param>
        ///// <returns>目标工作簿</returns>
        //public HSSFWorkbook addImages2Excel(HSSFWorkbook workbook, int sheetIndex, string[] imagePathList, string imgCellStr )
        //{
        //    string sheetName = workbook.GetSheetName(sheetIndex);
        //    return addImages2Excel(workbook, sheetName, imagePathList, imgCellStr);
        //}


        ///// <summary>
        ///// 添加一组图片到excel,各个图片间隔1行
        ///// </summary>
        ///// <param name="workbook">源工作簿</param>
        ///// <param name="sheetName">工作表sheet索引</param>
        ///// <param name="imagePathList">图片文件路径清单</param>
        ///// <param name="imgCellStr">图片位置标志字符串</param>
        ///// <returns>目标工作簿</returns>
        //public HSSFWorkbook addImages2Excel(HSSFWorkbook workbook, string sheetName, string[] imagePathList,  string imgCellStr )
        //{
        //    try
        //    {
        //        HSSFSheet sheet = (HSSFSheet)workbook.GetSheet(sheetName);
        //        int rowIndex = selectPosition(sheet, imgCellStr).X;//获取要插入的起始行号
        //        int colIndex = selectPosition(sheet, imgCellStr).Y;//获取要插入的起始列号
        //        for (int i = 0; i < imagePathList.Length; i++)
        //        {
        //            if (!File.Exists(imagePathList[i]))
        //            {
        //                continue;
        //            }
        //            #region 扩充行,并设置格式为上一行
        //            //先扩充一行
        //            sheet.ShiftRows(rowIndex + i+1,                                 //--开始行
        //                sheet.LastRowNum,                            //--结束行
        //                1,                             //--移动大小(行数)--往下移动
        //                true,                                   //是否复制行高
        //                false,                                  //是否重置行高
        //                true                                    //是否移动批注
        //                );
        //            // 对批量移动后空出的空行插，创建相应的行，并以插入行的上一行为格式源(即：InsertRowIndex-1的那一行)

        //            HSSFRow targetRow = null;
        //            HSSFCell sourceCell = null;
        //            HSSFCell targetCell = null;
        //            HSSFRow mySourceStyleRow = (HSSFRow)sheet.GetRow(rowIndex);
        //            if (mySourceStyleRow == null)
        //                continue;

        //            targetRow = (HSSFRow)sheet.CreateRow(rowIndex + i+1);//先新增一行
        //            targetRow = (HSSFRow)sheet.GetRow(rowIndex + i + 1);//先新增一行
        //            if (targetRow == null)
        //                continue;

        //            //复制格式到新的行
        //            for (int m = mySourceStyleRow.FirstCellNum; m < mySourceStyleRow.LastCellNum; m++)
        //            {
        //                sourceCell = (HSSFCell)mySourceStyleRow.GetCell(m);
        //                if (sourceCell == null)
        //                    continue;
        //                targetCell = (HSSFCell)targetRow.CreateCell(m);
        //                //targetCell.Encoding = sourceCell.Encoding;
        //                targetCell.CellStyle = sourceCell.CellStyle;
        //                targetCell.SetCellType(sourceCell.CellType);

        //            }
        //            #endregion

        //            HSSFRow temprow = (HSSFRow)sheet.GetRow(rowIndex + i);//递增的行
        //            if (temprow == null)
        //                continue;
        //            HSSFCell cell = (HSSFCell)temprow.GetCell(colIndex);
        //            cell.SetCellValue("");
        //            temprow.Height = 290 * 20;

        //            workbook = addImage2Excel(workbook, sheetName, imagePathList[i], colIndex, rowIndex + i);

        //        }
        //        return workbook;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex,"");
        //        return workbook;
        //    }
        //}

        /// <summary>
        /// 添加单个图片到指定sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="row1">插入位置:行号</param>
        /// <param name="col1">插入位置:列号</param>
        /// <returns></returns>
        private IWorkbook addImageToSheet(IWorkbook workbook, int sheetIndex, string imagePath, int row1, int col1)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                //sheet.PrintSetup.UsePage = true;
                //sheet.PrintSetup.PageStart = 5;
                int[,] range = getCellMergeArea(sheet, row1, col1);

                byte[] bytes = System.IO.File.ReadAllBytes(imagePath);
                int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);

                //HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                IDrawing patriarch = sheet.CreateDrawingPatriarch();

                //int lastCellWidth = sheet.GetColumnWidth(range[1, 1]);
                //int lastCellHeight = sheet.GetRow(range[1, 0]).Height + 0;
                //add a picture
                /*
                 关于HSSFClientAnchor(dx1,dy1,dx2,dy2,col1,row1,col2,row2)的参数，有必要在这里说明一下：
                 * dx1：起始单元格的x偏移量，表示直线起始位置距单元格左侧的距离；
                 * dy1：起始单元格的y偏移量，表示直线起始位置距单元格上侧的距离；
                 * dx2：终止单元格的x偏移量，表示直线起始位置距单元格左侧的距离；
                 * dy2：终止单元格的y偏移量，表示直线起始位置距单元格上侧的距离；
                 * col1：起始单元格列序号，从0开始计算；
                 * row1：起始单元格行序号，从0开始计算；
                 * col2：终止单元格列序号，从0开始计算；
                 * row2：终止单元格行序号，从0开始计算；
                 */
                IClientAnchor anchor = null;
                if (workbook is HSSFWorkbook)
                    anchor = new HSSFClientAnchor(0, 0, 0, 0, range[0, 1], range[0, 0], range[1, 1] + 1, range[1, 0] + 1);
                else if (workbook is XSSFWorkbook)
                    anchor = new XSSFClientAnchor(0, 0, 0, 0, range[0, 1], range[0, 0], range[1, 1] + 1, range[1, 0] + 1);

                //HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                IPicture pict = patriarch.CreatePicture(anchor, pictureIdx);

                //Resize(double scaleX, double scaleY),2个参数代表宽高缩放百分比
                //pict.Resize();//图片大小不拉伸
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }
        }

        /// <summary>
        /// 添加2个图片到2图片模板
        /// </summary>
        /// <param name="filePath">模板路径</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="imagePath1">图片1路径</param>
        /// <param name="imagePath2">图片2路径</param>
        /// <param name="imgFlag1">图片1标记</param>
        /// <param name="imgFlag2">图片2标记</param>
        /// <param name="endRowFlag">结束行标记</param>
        /// <param name="dArray">静态数据</param>
        public void addImgTo2ImgWorkbook(string filePath, int sheetIndex, 
            string imagePath1, string imagePath2, string imgFlag1, string imgFlag2, string endRowFlag, object[]dArray)
        {
            try
            {
                IWorkbook workbook = loadExcelWorkbookI(filePath);
                HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(sheetIndex);

                //添加第一张图片
                Point p = selectPosition(sheet, imgFlag1);
                workbook = addImageToSheet(workbook, sheetIndex, imagePath1, p.X, p.Y);

                //如果图片2的路径不为"",则填充图片
                if (!imagePath2.Equals(""))
                {
                    workbook = addImageToSheet(workbook, sheetIndex, imagePath2, selectPosition(sheet, imgFlag2).X, selectPosition(sheet, imgFlag2).Y);
                }
                //如果图片2的路径为"",则合并剩余单元格
                else
                {
                    //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                    //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(36, 39, 0, 7));
                    NPOI.SS.Util.CellRangeAddress region = new NPOI.SS.Util.CellRangeAddress(selectPosition(sheet, imgFlag2).X-1, selectPosition(sheet, endRowFlag).X,
                        selectPosition(sheet, endRowFlag).Y, getCellMergeArea(sheet, endRowFlag)[1, 1]);
                    sheet.AddMergedRegion(region);

                    #region 创建合并后单元格风格,黑边框,水平居中,垂直靠上
                    ICellStyle style = workbook.CreateCellStyle();
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BottomBorderColor = HSSFColor.Black.Index;
                    style.LeftBorderColor = HSSFColor.Black.Index;
                    style.RightBorderColor = HSSFColor.Black.Index;
                    style.TopBorderColor = HSSFColor.Black.Index;
                    style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//水平居中
                    style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;//垂直靠上
                    style.SetFont(HSSFCellUtil.GetRow(region.LastRow, sheet).GetCell(region.LastColumn).CellStyle.GetFont(workbook));
                    #endregion

                    for (int i = region.FirstRow; i <= region.LastRow; i++)
                    {
                        IRow row = HSSFCellUtil.GetRow(i, sheet);
                        for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                        {
                            ICell singleCell = HSSFCellUtil.GetCell(row, (short)j);
                            singleCell.SetCellValue("");
                            singleCell.CellStyle = style;
                        }
                    }
                    HSSFCellUtil.GetRow(region.FirstRow, sheet).GetCell(region.FirstColumn).SetCellValue(endRowFlag);
                    //sheet = replaceCellValue(sheet, endRowFlag, imgFlag2);
                }
                workbook = fillDataToExcelByValue(workbook, sheetIndex, dArray2Dictionary(dArray));
                saveExcelWithoutAsk(filePath, workbook);
            }
            catch (Exception ex)
            {
                WriteLog(ex,"");
            }

        }

        /// <summary>
        /// 添加2个图片到2图片模板
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="imagePath1">图片1路径</param>
        /// <param name="imagePath2">图片2路径</param>
        /// <param name="imgFlag1">图片1标记</param>
        /// <param name="imgFlag2">图片2标记</param>
        /// <param name="endRowFlag">结束行标记</param>
        /// <param name="dArray">静态数据</param>
        public IWorkbook addImgTo2ImgWorkbook(IWorkbook workbook, int sheetIndex,
            string imagePath1, string imagePath2, string imgFlag1, string imgFlag2, string endRowFlag, object[] dArray)
        {
            try
            {
                ISheet sheet = (ISheet)workbook.GetSheetAt(sheetIndex);

                //添加第一张图片
                Point p = selectPosition(sheet, imgFlag1);
                workbook = addImageToSheet(workbook, sheetIndex, imagePath1, p.X, p.Y);

                //如果图片2的路径不为"",则填充图片
                if (!imagePath2.Equals(""))
                {
                    workbook = addImageToSheet(workbook, sheetIndex, imagePath2, selectPosition(sheet, imgFlag2).X, selectPosition(sheet, imgFlag2).Y);
                }
                //如果图片2的路径为"",则合并剩余单元格
                else
                {
                    //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                    //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(36, 39, 0, 7));
                    NPOI.SS.Util.CellRangeAddress region = new NPOI.SS.Util.CellRangeAddress(selectPosition(sheet, imgFlag2).X - 1, selectPosition(sheet, endRowFlag).X,
                        selectPosition(sheet, endRowFlag).Y, getCellMergeArea(sheet, endRowFlag)[1, 1]);
                    sheet.AddMergedRegion(region);

                    #region 创建合并后单元格风格,黑边框,水平居中,垂直靠上
                    ICellStyle style = workbook.CreateCellStyle();
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BottomBorderColor = HSSFColor.Black.Index;
                    style.LeftBorderColor = HSSFColor.Black.Index;
                    style.RightBorderColor = HSSFColor.Black.Index;
                    style.TopBorderColor = HSSFColor.Black.Index;
                    style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//水平居中
                    style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;//垂直靠上
                    style.SetFont(CellUtil.GetRow(region.LastRow, sheet).GetCell(region.LastColumn).CellStyle.GetFont(workbook));
                    #endregion

                    for (int i = region.FirstRow; i <= region.LastRow; i++)
                    {
                        IRow row = CellUtil.GetRow(i, sheet);
                        for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                        {
                            ICell singleCell = HSSFCellUtil.GetCell(row, (short)j);
                            singleCell.SetCellValue("");
                            singleCell.CellStyle = style;
                        }
                    }
                    CellUtil.GetRow(region.FirstRow, sheet).GetCell(region.FirstColumn).SetCellValue(endRowFlag);
                    //sheet = replaceCellValue(sheet, endRowFlag, imgFlag2);
                }
                workbook = fillDataToExcelByValue(workbook, sheetIndex, dArray2Dictionary(dArray));
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }

        }

        /// <summary>
        /// 添加1个图片到1图片模板
        /// </summary>
        /// <param name="filePath">模板路径</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="imagePath1">图片1路径</param>
        /// <param name="imgFlag1">图片1标记</param>
        /// <param name="dArray">静态数据</param>
        public void addImgTo1ImgWorkbook(string filePath, int sheetIndex,
            string imagePath1, string imgFlag1, object[] dArray)
        {
            try
            {
                IWorkbook workbook = loadExcelWorkbookI(filePath);
                HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(sheetIndex);

                //添加第一张图片
                Point p = selectPosition(sheet, imgFlag1);
                workbook = addImageToSheet(workbook, sheetIndex, imagePath1, p.X, p.Y);
                
                workbook = fillDataToExcelByValue(workbook, sheetIndex, dArray2Dictionary(dArray));
                saveExcelWithoutAsk(filePath, workbook);
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
            }

        }

        /// <summary>
        /// 添加1个图片到1图片模板
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="imagePath1">图片1路径</param>
        /// <param name="imgFlag1">图片1标记</param>
        /// <param name="dArray">静态数据</param>
        public IWorkbook addImgTo1ImgWorkbook(IWorkbook workbook, int sheetIndex, string imagePath1, string imgFlag1, object[] dArray)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);

                //添加第一张图片
                Point p = selectPosition(sheet, imgFlag1);
                workbook = addImageToSheet(workbook, sheetIndex, imagePath1, p.X, p.Y);
                workbook = fillDataToExcelByValue(workbook, sheetIndex, dArray2Dictionary(dArray));
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }

        }


        /// <summary>
        /// 添加图片到指定excel,图片指定大小
        /// </summary>
        /// <param name="workbook">源工作簿</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="col1">列号</param>
        /// <param name="row1">行号</param>
        /// <returns>目标工作簿</returns>
        public HSSFWorkbook addImage2Excel(HSSFWorkbook workbook, int sheetIndex, string imagePath, int col1, int row1 )
        {
            string sheetName = workbook.GetSheetName(sheetIndex);
            return addImage2Excel( workbook, sheetName, imagePath, col1, row1 );

        }

        /// <summary>
        /// 添加图片到指定excel,图片指定大小
        /// </summary>
        /// <param name="workbook">源工作簿</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="col1">列号</param>
        /// <param name="row1">行号</param>
        /// <returns>目标工作簿</returns>
        public HSSFWorkbook addImage2Excel(HSSFWorkbook workbook, string sheetName, string imagePath, int col1, int row1)
        {
            try
            {
                #region
                ////add picture data to this workbook.
                //string imagePath = "";
                //OpenFileDialog openDialog = new OpenFileDialog();
                //openDialog.DefaultExt = "jpg";
                //openDialog.Filter = "图片文件|*.jpg";
                //MemoryStream ms = new MemoryStream();
                //if (openDialog.ShowDialog() == DialogResult.OK)
                //{
                //    imagePath = openDialog.FileName;
                //}
                //else
                //{
                //    return;
                //}

                //imagePath = saveResizedPhoto(PhotoSizeChange(imagePath, width, height), "D:\\test.jpg");
                #endregion

                
                byte[] bytes = System.IO.File.ReadAllBytes(imagePath);
                int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);

                // Create the drawing patriarch.  This is the top level container for all shapes.
                HSSFSheet sheet = (HSSFSheet)workbook.GetSheet(sheetName);
                HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

                //add a picture
                /*
                 关于HSSFClientAnchor(dx1,dy1,dx2,dy2,col1,row1,col2,row2)的参数，有必要在这里说明一下：
                 * dx1：起始单元格的x偏移量，表示直线起始位置距单元格左侧的距离；
                 * dy1：起始单元格的y偏移量，表示直线起始位置距单元格上侧的距离；
                 * dx2：终止单元格的x偏移量，表示直线起始位置距单元格左侧的距离；
                 * dy2：终止单元格的y偏移量，表示直线起始位置距单元格上侧的距离；
                 * col1：起始单元格列序号，从0开始计算；
                 * row1：起始单元格行序号，从0开始计算；
                 * col2：终止单元格列序号，从0开始计算；
                 * row2：终止单元格行序号，从0开始计算；
                 */
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, col1, row1, col1+1, row1+1);
                HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

                //Resize(double scaleX, double scaleY),2个参数代表宽高缩放百分比
                //pict.Resize();//图片大小不拉伸
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="photo">要保存的图片</param>
        /// <param name="strPhoto">保存路径</param>
        /// <returns>保存路径</returns>
        private string saveResizedPhoto(System.Drawing.Image photo, string strPhoto)
        {
            System.IO.FileStream ms = new System.IO.FileStream(strPhoto, FileMode.Create);
            photo.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            ms.Close();
            return strPhoto;
        }

        /// <summary>
        /// 处理图片大小到指定尺寸，返回值为一个Image对象，使用Image对象的Save方法就可以保存该图片
        /// </summary>
        /// <param name="strPhoto">保存路径</param>
        /// <param name="width">图片宽度,像素</param>
        /// <param name="height">图片高度,像素</param>
        /// <returns>图片对象</returns>
        private System.Drawing.Image PhotoSizeChange(string strPhoto, int width, int height)
        {
            //strPhoto是原来的图片文件所在的物理路径
            //处理图片功能
            System.Drawing.Image image = new Bitmap(strPhoto);//得到原图
            //创建指定大小的图
            System.Drawing.Image newImage = image.GetThumbnailImage(width, height, null, new IntPtr());
            Graphics g = Graphics.FromImage(newImage);
            //将原图画到指定的图上
            g.DrawImage(newImage, 0, 0, newImage.Width, newImage.Height);
            g.Dispose();
            return newImage;
        }

        /// <summary>
        /// 获取图片的宽和高
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>Point对象[Width:宽,Height:高]</returns>
        private Point getImageSize(string filePath)
        {
            Point p = new Point();
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            System.Drawing.Image  tempimage = System.Drawing.Image.FromStream(fs, true);           
            p.X = tempimage.Width;//宽
            p.Y = tempimage.Height;//高
            fs.Close();
            return p;
        }

        /// <summary>
        /// 获取图片的水平和垂直分辨率(以 像素/英寸 为单位)
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>2个元素的数组,[0]:垂直分辨率,[1]:水平分辨率</returns>
        private double[] getImageResolution(string filePath){

            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            System.Drawing.Image tempimage = System.Drawing.Image.FromStream(fs, true);
            double[] d = new double[2];
            d[0] = tempimage.VerticalResolution;//垂直分辨率
            d[1] = tempimage.HorizontalResolution;//水平分辨率
            fs.Close();
            return d;
        }

        /// <summary>
        /// 检查图片是否符合尺寸要求(按像素)
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="width">像素宽度</param>
        /// <param name="height">像素高度</param>
        /// <returns>是否合格</returns>
        public Boolean isDimensionsMatch(string filePath,int width, int height)
        {
            Point p = getImageSize(filePath);
            if (p.X == width && p.Y == height)
            {
                return true;
            }
            else
            {
                return false;

            }
        }

        /// <summary>
        /// 读取指定xls文件的指定sheet
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheeIndex">工作表sheet索引名</param>
        /// <returns>读取到的工作表,失败返回空</returns>
        public ISheet loadExcelSheetI(string filePath, int sheeIndex)
        {
            try
            {
                FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                IWorkbook wb = loadExcelWorkbookI(filePath);
                ISheet sheet = wb.GetSheetAt(sheeIndex);
                file.Close();
                return sheet;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }


        /// <summary>
        /// 读取指定xls文件为workbook
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>读取到的工作表,失败返回空</returns>
        public IWorkbook loadExcelWorkbookI(string filePath)
        {
            try
            {
                FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                IWorkbook wb = new HSSFWorkbook(file);
                file.Close();
                file.Dispose();
                return wb;
            }
            catch 
            {
                try
                {
                    FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    IWorkbook wb = new XSSFWorkbook(file);
                    file.Close();
                    file.Dispose();
                    return wb;
                }catch (Exception ex)
                {
                    WriteLog(ex, "");
                    return null;
                }
            }
        }

        /// <summary>
        /// 在指定位置插入行
        /// </summary>
        /// <param name="workbook">源workbook工作簿</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="InsertRowIndex">插入行位置,行号</param>
        /// <param name="mySourceStyleRow">要插入的行</param>
        /// <returns>目标工作簿</returns>
        public IWorkbook MyInsertRow(IWorkbook workbook, int sheetIndex, int InsertRowIndex, List<IRow> mySourceStyleRow)
        {
            string sheetName = workbook.GetSheetName(sheetIndex);
            return MyInsertRow(workbook, sheetName, InsertRowIndex, mySourceStyleRow);
        }

        
        /// <summary>
        /// 在指定位置插入行
        /// </summary>
        /// <param name="wb">源workbook工作簿</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="InsertRowIndex">插入行位置,行号</param>
        /// <param name="mySourceStyleRow">要插入的行</param>
        /// <returns>目标工作簿</returns>
        public IWorkbook MyInsertRow(IWorkbook workbook, string sheetName, int InsertRowIndex, List<IRow> mySourceStyleRow)
        {
            ISheet sheet = workbook.GetSheet(sheetName);
            int InsertRowCount = mySourceStyleRow.Count;
            #region 批量移动行
            sheet.ShiftRows(
                InsertRowIndex,                                 //--开始行
                sheet.LastRowNum,                            //--结束行
                InsertRowCount,                             //--移动大小(行数)--往下移动
                true,                                   //是否复制行高
                false                                  //是否重置行高
                );
            #endregion

            #region 对批量移动后空出的空行插，创建相应的行，并以插入行的上一行为格式源(即：InsertRowIndex-1的那一行)
            for (int i = InsertRowIndex; i < InsertRowIndex + InsertRowCount - 1; i++)
            {
                IRow targetRow = null;
                ICell sourceCell = null;
                ICell targetCell = null;

                targetRow = sheet.CreateRow(i + 1);

                for (int m = mySourceStyleRow[i - InsertRowIndex].FirstCellNum; m < mySourceStyleRow[i - InsertRowIndex].LastCellNum; m++)
                {
                    sourceCell = mySourceStyleRow[i - InsertRowIndex].GetCell(m);
                    if (sourceCell == null)
                        continue;
                    targetCell = targetRow.CreateCell(m);
                    //targetCell.Encoding = sourceCell.Encoding;
                    targetCell.CellStyle = sourceCell.CellStyle;
                    targetCell.SetCellType(sourceCell.CellType);

                }
                //CopyRow(sourceRow, targetRow);

                //Util.CopyRow(sheet, sourceRow, targetRow);
            }

            IRow firstTargetRow = null;
            ICell firstSourceCell = null;
            ICell firstTargetCell = null;
            for (int i = 0; i < InsertRowCount; i++)
            {
                firstTargetRow = sheet.GetRow(i + InsertRowIndex);
                if (firstTargetRow == null)
                    continue;
                for (int m = mySourceStyleRow[i].FirstCellNum; m < mySourceStyleRow[i].LastCellNum; m++)
                {
                    firstSourceCell = mySourceStyleRow[i].GetCell(m);
                    if (firstSourceCell == null)
                        continue;
                    firstTargetCell = firstTargetRow.CreateCell(m);

                    //firstTargetCell.Encoding = firstSourceCell.Encoding;
                    firstTargetCell.CellStyle.CloneStyleFrom(firstSourceCell.CellStyle);
                    firstTargetCell.SetCellType(firstSourceCell.CellType);

                    //写值
                    copyCellValue(firstTargetCell, firstSourceCell);
                }
            }
            #endregion

            return workbook;
        }

        
        /// <summary>
        /// 单元格值赋值
        /// </summary>
        /// <param name="targetCell">目标单元格</param>
        /// <param name="sourceCell">源单元格</param>
        public void copyCellValue(ICell targetCell, ICell sourceCell)
        {

            string ValueType = sourceCell.CellType.ToString();
            switch (ValueType)
            {
                case "String"://字符串型
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                    break;
                case "Numeric"://数值型,包括浮点和整数
                    targetCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case "DateTime"://布尔型
                    System.DateTime dateV;
                    System.DateTime.TryParse(sourceCell.DateCellValue.ToString(), out dateV);
                    targetCell.SetCellValue(dateV);
                    break;
                case "Boolean"://布尔型
                    bool boolV = false;
                    bool.TryParse(sourceCell.BooleanCellValue.ToString(), out boolV);
                    targetCell.SetCellValue(boolV);
                    break;
                case "IRichTextString"://RichStirng类型
                    targetCell.SetCellValue(sourceCell.RichStringCellValue);
                    break;
                default:
                    targetCell.SetCellValue("");
                    break;
            }
        }

        /// 获取sheet的最小列号和最大列号
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <returns>[0]:最小列号,[1]:最大列号</returns>
        public static int[] getColumnRange(ISheet sheet)
        {
            int[] p = { 0, 0 };
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                if (row.FirstCellNum < p[0])
                {
                    p[0] = row.FirstCellNum;
                }
                if (row.LastCellNum > p[1])
                {
                    p[1] = row.LastCellNum;
                }
            }
            p[0] = p[0] < 0 ? 0 : p[0];
            return p;
        }

        
        #region 废弃使用"C12"作为单元格坐标
        ////填写数据到excel,根据单元格坐标
        ////Dictionary的key为坐标(如"C3"),value为值
        ///// <summary>
        ///// 填写数据到excel,根据单元格坐标.Dictionary的key为坐标(如"C3"),value为值
        ///// </summary>
        ///// <param name="sourceExcelPath">源工作簿文件路径</param>
        ///// <param name="targetExcelPath">目标工作簿文件路径</param>
        ///// <param name="map">Dictionary数据对象</param>
        ///// <param name="sheetIndex">工作表sheet索引</param>
        //public void fillDataToExcel( string sourceExcelPath, string targetExcelPath, Dictionary<string, string> map, int sheetIndex)
        //{
        //    string sheetName = loadExcelWorkbook(sourceExcelPath).GetSheetName(sheetIndex);
        //    fillDataToExcel(sourceExcelPath, targetExcelPath, map, sheetName);
        //}

        ////填写数据到excel,根据单元格坐标
        ////Dictionary的key为坐标(如"C3"),value为值
        ///// <summary>
        ///// 填写数据到excel,根据单元格坐标.Dictionary的key为坐标(如"C3"),value为值
        ///// </summary>
        ///// <param name="sourceExcelPath">源工作簿文件路径</param>
        ///// <param name="targetExcelPath">目标工作簿文件路径</param>
        ///// <param name="map">Dictionary数据对象</param>
        ///// <param name="sheetName">工作表名</param>
        //public void fillDataToExcel( string sourceExcelPath, string targetExcelPath,  Dictionary<string, string> map, string sheetName)
        //{
        //    try
        //    {
        //        HSSFWorkbook wb = loadExcelWorkbook(sourceExcelPath);
        //        HSSFSheet sheet = (HSSFSheet)wb.GetSheet(sheetName);//获取模板sheet

        //        foreach (var oneMapPoint in map)
        //        {
        //            string key = oneMapPoint.Key.ToString();
        //            string value = oneMapPoint.Value.ToString();
        //            int row = disassemblyToNumber(key) - 1;
        //            int col = System26ToNumber(disassemblyToString(key)) - 1;
        //            sheet.GetRow(row).GetCell(col).SetCellValue(value);
        //        }
        //        saveExcelWithoutAsk(targetExcelPath, wb);
        //        return;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return;
        //    }
        //}
        #endregion

        /// <summary>
        /// 填写数据到excel,根据单元格值,如有重复值,则选择[行,列]靠前的值.Dictionary的key代表被替换的文本,value代表用于替换的文本
        /// </summary>
        /// <param name="sourceExcelPath">源工作簿文件路径</param>
        /// <param name="targetExcelPath">目标工作簿文件路径</param>
        /// <param name="map">Dictionary数据对象</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        public void fillDataToExcelByValue( string sourceExcelPath, string targetExcelPath, Dictionary<string, string> map, int sheetIndex )
        {
            string sheetName = loadExcelWorkbookI(sourceExcelPath).GetSheetName(sheetIndex);
            fillDataToExcelByValue(sourceExcelPath, targetExcelPath, map, sheetName);
        }

        /// <summary>
        /// 填写数据到excel,根据单元格值,如有重复值,则选择[行,列]靠前的值.Dictionary的key代表被替换的文本,value代表用于替换的文本
        /// </summary>
        /// <param name="sourceExcelPath">源工作簿文件路径</param>
        /// <param name="targetExcelPath">目标工作簿文件路径</param>
        /// <param name="map">Dictionary数据对象</param>
        /// <param name="sheetName">工作表名</param>
        public void fillDataToExcelByValue( string sourceExcelPath, string targetExcelPath, Dictionary<string, string> map, string sheetName )
        {
            try
            {
                IWorkbook wb = loadExcelWorkbookI(sourceExcelPath);
                ISheet sheet = wb.GetSheet(sheetName);//获取模板sheet

                foreach (var oneMapPoint in map)
                {
                    string key = oneMapPoint.Key.ToString();
                    string value = oneMapPoint.Value.ToString();
                    replaceCellValue(sheet, value, key);
                }
                saveExcelWithoutAsk(targetExcelPath, wb);
                return;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return;
            }
        }

        //Dictionary的key代表被替换的文本,value代表用于替换的文本
        /// <summary>
        /// 填写数据到workbook,根据单元格值,如有重复值,则选择[行,列]靠前的值.Dictionary的key代表被替换的文本,value代表用于替换的文本
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="map">Dictionary数据对象</param>
        /// <returns>workbook</returns>
        private IWorkbook fillDataToExcelByValue(IWorkbook workbook, int sheetIndex, Dictionary<string, string> map)
        {
            try
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);//获取模板sheet

                foreach (var oneMapPoint in map)
                {
                    string key = oneMapPoint.Key.ToString();
                    string value = oneMapPoint.Value.ToString();
                    replaceCellValue(sheet, value, key);
                }
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }
        }

        //调换sheet的行列,行列转换
        /// <summary>
        /// 调换sheet的行列,行列转换
        /// </summary>
        /// <param name="wb">源工作簿</param>
        /// <param name="sourceSheetName">源工作表名</param>
        /// <param name="targetSheetName">目标工作表名</param>
        /// <returns>目标工作表</returns>
        public HSSFWorkbook changeRowAndColumn(HSSFWorkbook wb, string sourceSheetName, string targetSheetName)
        {
            if (sourceSheetName.ToUpper().Equals(targetSheetName.ToUpper()))
                return wb;
            HSSFSheet sheet = (HSSFSheet)wb.GetSheet(sourceSheetName);
            if (sheet == null)
                return wb;
            HSSFSheet sheet2 = (HSSFSheet)wb.CreateSheet(targetSheetName);

            //先添加所有单元格
            for (int col = 0; col <= getColumnRange(sheet)[1]; col++)
            {
                HSSFRow dataRow = (HSSFRow)sheet2.CreateRow(col);
                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    HSSFCell targetCell = (HSSFCell)dataRow.CreateCell(row);
                    HSSFCell sourceCell = (HSSFCell)sheet.GetRow(row).GetCell(col);
                    if (sourceCell == null)
                        continue;
                    copyCellValue(targetCell, sourceCell);
                }
            }
            return wb;

        }
        
        /// <summary>
        /// 创建新的sheet,如果目标索引存在,则此sheet重命名为目标sheet名,不存在则创建
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetIndex">目标sheet索引</param>
        /// <param name="sheetName">sheet名</param>
        /// <returns>新增后的workbook</returns>
        public IWorkbook addNewSheet(IWorkbook workbook, int sheetIndex, string sheetName)
        {
            ISheet toSheet;
            try
            {
                try
                {
                    toSheet = (ISheet)workbook.GetSheetAt(sheetIndex);
                    workbook.SetSheetName(sheetIndex, sheetName);
                }
                catch
                {
                    workbook.CreateSheet(sheetName);
                    toSheet = (ISheet)workbook.GetSheetAt(sheetIndex);
                }
                return workbook;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return workbook;
            }

        }

        ////添加一个sheet
        ///// <summary>
        ///// 添加一个sheet到工作簿,默认Sheet名为Sheet1,下一个为Sheet2
        ///// </summary>
        ///// <param name="workbook">目标工作簿</param>
        ///// <param name="wbPaths">源工作簿文件路径</param>
        ///// <returns>合并后的工作簿</returns>
        //public HSSFWorkbook addSheetsToWorkbook(HSSFWorkbook workbook, object[] wbPaths)
        //{
        //    try
        //    {
        //        for (int i = 0; i < wbPaths.Length; i++)
        //        {
        //            HSSFSheet sheet = loadExcelSheet((string)wbPaths[i], 0);
        //            if (sheet == null)
        //                return workbook;
        //            addNewSheet(workbook, i + 1, "Sheet" + (i + 1), sheet);
        //            CopySheet(workbook, sheet, workbook.GetSheetAt(i + 1), true);
        //        }
        //        return workbook;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex,"");
        //        return workbook;
        //    }
        //}

        ////合并多个workbook到指定Workbook
        ///// <summary>
        ///// 合并多个workbook到指定Workbook
        ///// </summary>
        ///// <param name="fromWorkbook">源工作簿路径</param>
        ///// <param name="wbPaths">源工作簿文件路径</param>
        ///// <param name="toWorkbook">目标工作簿路径</param>
        ///// <returns>成功则犯规目标工作簿,失败则返回源工作簿</returns>
        //public string addWorkbooks(string fromWorkbook, object[] wbPaths, string toWorkbook )
        //{
        //    try
        //    {
        //        HSSFWorkbook wb = loadExcelWorkbook(fromWorkbook);
        //        wb = addSheetsToWorkbook(wb, wbPaths);
        //        saveExcelWithoutAsk(toWorkbook, wb);
        //        return toWorkbook;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex,"");
        //        return fromWorkbook;
        //    }
        //}

        #region 网上的复制功能,参见:http://blog.csdn.net/wutbiao/article/details/8696446

        //复制一个单元格样式到目的单元格样式 
        /// <summary>
        /// 复制一个单元格样式到目的单元格样式 
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="fromStyle"></param>
        /// <param name="toStyle"></param>
        public static ICellStyle CopyCellStyle(IWorkbook wb, ICellStyle fromStyle, ICellStyle toStyle)
        {
            //使用异常处理,忽略不能识别的单元格风格
            try
            {
                toStyle.Alignment = fromStyle.Alignment;
                //边框和边框颜色
                toStyle.BorderBottom = fromStyle.BorderBottom;
                toStyle.BorderLeft = fromStyle.BorderLeft;
                toStyle.BorderRight = fromStyle.BorderRight;
                toStyle.BorderTop = fromStyle.BorderTop;
                toStyle.TopBorderColor = fromStyle.TopBorderColor;
                toStyle.BottomBorderColor = fromStyle.BottomBorderColor;
                toStyle.RightBorderColor = fromStyle.RightBorderColor;
                toStyle.LeftBorderColor = fromStyle.LeftBorderColor;
                //背景和前景
                toStyle.FillBackgroundColor = fromStyle.FillBackgroundColor;
                toStyle.FillForegroundColor = fromStyle.FillForegroundColor;
                toStyle.DataFormat = fromStyle.DataFormat;
                toStyle.FillPattern = fromStyle.FillPattern;
                //toStyle.Hidden=fromStyle.Hidden;
                toStyle.IsHidden = fromStyle.IsHidden;
                toStyle.Indention = fromStyle.Indention;//首行缩进
                toStyle.IsLocked = fromStyle.IsLocked;
                toStyle.Rotation = fromStyle.Rotation;//旋转
                toStyle.VerticalAlignment = fromStyle.VerticalAlignment;
                toStyle.WrapText = fromStyle.WrapText;
                toStyle.SetFont(fromStyle.GetFont(wb));
                return toStyle;
            }
            catch(Exception ex)
            {
                //WriteLog(ex, "");
                System.Console.WriteLine(ex.ToString());
                return toStyle;
            }
        }

        //复制打印设置
        /// <summary>
        /// 复制打印设置
        /// </summary>
        /// <param name="fromSheet"></param>
        /// <param name="toSheet"></param>
        public static void CopyPrintSetup(ISheet fromSheet, ISheet toSheet)
        {
            //打印设置
            toSheet.PrintSetup.Landscape = fromSheet.PrintSetup.Landscape;
            toSheet.PrintSetup.FitWidth = fromSheet.PrintSetup.FitWidth;
            toSheet.PrintSetup.FitHeight = fromSheet.PrintSetup.FitHeight;
            toSheet.PrintSetup.PaperSize = fromSheet.PrintSetup.PaperSize;
            toSheet.PrintSetup.UsePage = fromSheet.PrintSetup.UsePage;
            toSheet.PrintSetup.PageStart = fromSheet.PrintSetup.PageStart;
            toSheet.IsPrintGridlines = fromSheet.IsPrintGridlines;
            toSheet.PrintSetup.NoColor = fromSheet.PrintSetup.NoColor;
            toSheet.PrintSetup.Draft = fromSheet.PrintSetup.Draft;
            toSheet.PrintSetup.LeftToRight = fromSheet.PrintSetup.LeftToRight;
            toSheet.PrintSetup.CellError = fromSheet.PrintSetup.CellError;
            toSheet.PrintSetup.Notes = fromSheet.PrintSetup.Notes;
            toSheet.Header.Center = fromSheet.Header.Center;
            toSheet.Header.Left = fromSheet.Header.Left;
            toSheet.Header.Right = fromSheet.Header.Right;
            toSheet.Footer.Center = fromSheet.Footer.Center;
            toSheet.Footer.Left = fromSheet.Footer.Left;
            toSheet.Footer.Right = fromSheet.Footer.Right;
            toSheet.RepeatingRows = fromSheet.RepeatingRows;
            toSheet.RepeatingColumns = fromSheet.RepeatingColumns;
            toSheet.SetMargin(MarginType.BottomMargin, fromSheet.GetMargin(MarginType.BottomMargin));
            toSheet.SetMargin(MarginType.FooterMargin, fromSheet.GetMargin(MarginType.FooterMargin));
            toSheet.SetMargin(MarginType.HeaderMargin, fromSheet.GetMargin(MarginType.HeaderMargin));
            toSheet.SetMargin(MarginType.LeftMargin, fromSheet.GetMargin(MarginType.LeftMargin));
            toSheet.SetMargin(MarginType.RightMargin, fromSheet.GetMargin(MarginType.RightMargin));
            toSheet.SetMargin(MarginType.TopMargin, fromSheet.GetMargin(MarginType.TopMargin));
        }

        //复制表
        /// <summary>
        /// 复制表
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="fromSheet"></param>
        /// <param name="toSheet"></param>
        /// <param name="copyValueFlag"></param>
        public static void CopySheet(IWorkbook wb, ISheet fromSheet, ISheet toSheet, bool copyValueFlag)
        {
            //宽度
            for (int i = getColumnRange(fromSheet)[0]; i < getColumnRange(fromSheet)[1]; i++)
            {
                toSheet.SetColumnWidth(i, fromSheet.GetColumnWidth(i));
            }

            //合并区域处理
            MergerRegion(fromSheet, toSheet);
            //复制打印设置
            CopyPrintSetup(fromSheet, toSheet);

            //System.Collections.IEnumerator rows = fromSheet.GetRowEnumerator();
            //while (rows.MoveNext())
            //{
            //    IRow row = null;
            //    if (rows.Current is HSSFRow)
            //        row = rows.Current as HSSFRow;
            //    else
            //        row = rows.Current as XSSFRow;
            //    IRow newRow = toSheet.CreateRow(row.RowNum);
            //    CopyRow(wb, row, newRow, copyValueFlag);
            //}            

            int rowCount = fromSheet.LastRowNum;
            for (int i = 0; i <= rowCount; i++)
            {
                if (fromSheet.GetRow(i) == null)
                {
                    continue;
                }
                IRow row = null;
                if (fromSheet is HSSFSheet)
                    row = fromSheet.GetRow(i) as HSSFRow;
                else
                    row = fromSheet.GetRow(i) as XSSFRow;
                IRow newRow = toSheet.CreateRow(i);
                CopyRow(wb, row, newRow, copyValueFlag);

            }
        }

        /// <summary>
        /// 添加sheet到workbook的指定sheet
        /// </summary>
        /// <param name="wb">目标workbook</param>
        /// <param name="fromSheet">源sheet</param>
        /// <param name="sheetIndex">目标身体索引</param>
        /// <param name="copyValueFlag">是否复制值</param>
        /// <returns>新workbook</returns>
        public IWorkbook CopySheet(IWorkbook wb, ISheet fromSheet, int sheetIndex, bool copyValueFlag)
        {
            try
            {
                CopySheet(wb, fromSheet, wb.GetSheetAt(sheetIndex), copyValueFlag);
                return wb;
            }catch(Exception ex)
            {
                WriteLog(ex, "");
                return wb;
            }
        }

        /// <summary>
        /// 复制行
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="fromRow"></param>
        /// <param name="toRow"></param>
        /// <param name="copyValueFlag"></param>
        public static void CopyRow(IWorkbook wb, IRow fromRow, IRow toRow, bool copyValueFlag)
        {
            //System.Collections.IEnumerator cells = fromRow.GetEnumerator();//.GetRowEnumerator();
            toRow.Height = fromRow.Height;
            //判断是否隐藏,因为隐藏行的行高为取消隐藏后的行高
            if (fromRow.ZeroHeight == true)
            {
                toRow.ZeroHeight = true;
            }
            //while (cells.MoveNext())
            //{
            //    ICell cell = null;
            //    if (cells.Current is HSSFCell)
            //        cell = cells.Current as HSSFCell;
            //    else
            //        cell = cells.Current as XSSFCell;
            //    ICell newCell = toRow.CreateCell(cell.ColumnIndex);
            //    CopyCell(wb, cell, newCell, copyValueFlag);

            //}
            //int firstCount = fromRow.FirstCellNum;
            int cellCount = fromRow.LastCellNum;
            for (int i=0; i< cellCount; i++)
            {
                if(fromRow.GetCell(i) == null)
                {
                    continue;
                }
                ICell cell = null;
                if (fromRow is HSSFRow)
                    cell = fromRow.GetCell(i) as HSSFCell;
                else
                    cell = fromRow.GetCell(i) as XSSFCell;
                ICell newCell = toRow.CreateCell(i);
                CopyCell(wb, cell, newCell, copyValueFlag);

            }
        }
        
        /// <summary>
        /// 复制原有sheet的合并单元格到新创建的sheet
        /// </summary>
        /// <param name="fromSheet"></param>
        /// <param name="toSheet"></param>
        public static void MergerRegion(ISheet fromSheet, ISheet toSheet)
        {
            int sheetMergerCount = fromSheet.NumMergedRegions;
            for (int i = 0; i < sheetMergerCount; i++)
            {
                //Region mergedRegionAt = fromSheet.GetMergedRegion(i); //.MergedRegionAt(i);
                //CellRangeAddress[] cra = new CellRangeAddress[1];
                //cra[0] = fromSheet.GetMergedRegion(i);
                //Region[] rg = Region.ConvertCellRangesToRegions(cra);
                toSheet.AddMergedRegion(fromSheet.GetMergedRegion(i));
            }
        }
        
        /// <summary>
        /// 复制单元格
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="srcCell"></param>
        /// <param name="distCell"></param>
        /// <param name="copyValueFlag">true则连同cell的内容一起复制 </param>
        public static void CopyCell(IWorkbook wb, ICell srcCell, ICell distCell, bool copyValueFlag)
        {
            ICellStyle newstyle = wb.CreateCellStyle();
            //WriteLog(distCell.RowIndex+":"+distCell.ColumnIndex, "");
            newstyle = CopyCellStyle(wb, srcCell.CellStyle, newstyle);
            //样式
            distCell.CellStyle = newstyle;
            //distCell.CellStyle.CloneStyleFrom(newstyle);
            //评论
            if (srcCell.CellComment != null)
            {
                distCell.CellComment = srcCell.CellComment;
            }
            // 不同数据类型处理
            CellType srcCellType = srcCell.CellType;
            distCell.SetCellType(srcCellType);
            if (copyValueFlag)
            {
                if (srcCellType == CellType.Numeric)
                {
                    if (HSSFDateUtil.IsCellDateFormatted(srcCell))
                    {
                        distCell.SetCellValue(srcCell.DateCellValue);
                    }
                    else
                    {
                        distCell.SetCellValue(srcCell.NumericCellValue);
                    }
                }
                else if (srcCellType == CellType.String)
                {
                    distCell.SetCellValue(srcCell.RichStringCellValue);
                }
                else if (srcCellType == CellType.Blank)
                {
                    // nothing21
                }
                else if (srcCellType == CellType.Boolean)
                {
                    distCell.SetCellValue(srcCell.BooleanCellValue);
                }
                else if (srcCellType == CellType.Error)
                {
                    distCell.SetCellErrorValue(srcCell.ErrorCellValue);
                }
                else if (srcCellType == CellType.Formula)
                {
                    distCell.SetCellFormula(srcCell.CellFormula);
                }
                else
                { // nothing29
                }
            }
        }

        #endregion

        #region 数据类型转换模块

        //10进制(1-26)转A~Z
        /// <summary>
        /// 将指定的自然数转换为26进制表示。映射关系：[1-26] ->[A-Z]。
        /// </summary>
        /// <param name="n">自然数（如果无效，则返回空字符串）。</param>
        /// <returns>26进制表示。</returns>
        public static string NumberToSystem26(int n)
        {
            string s = string.Empty;
            while (n > 0)
            {
                int m = n % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                n = (n - m) / 26;
            }
            return s;
        }

        //A~Z转10进制
        /// <summary>
        /// 将指定的26进制表示转换为自然数。映射关系：[A-Z] ->[1-26]。
        /// </summary>
        /// <param name="s">26进制表示（如果无效，则返回0）。</param>
        /// <returns>自然数。</returns>
        public static int System26ToNumber(string s)
        {
            try
            {
                if (string.IsNullOrEmpty(s)) return 0;
                int n = 0;
                for (int i = s.Length - 1, j = 1; i >= 0; i--, j *= 26)
                {
                    char c = Char.ToUpper(s[i]);
                    if (c < 'A' || c > 'Z') return 0;
                    n += ((int)c - 64) * j;
                }
                return n;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return -1;
            }
        }

        /// <summary>
        /// json to Dictionary<string, object>
        /// </summary>
        /// <param name="jsonData"></param>
        /// <returns></returns>
        public static Dictionary<string, object> json2Dictionary(string jsonData)
        {

            Dictionary<string, string> dic = new Dictionary<string, string>();
            //实例化JavaScriptSerializer类的新实例
            JavaScriptSerializer jss = new JavaScriptSerializer();
            try
            {
                //将指定的 JSON 字符串转换为 Dictionary<string, object> 类型的对象
                return jss.Deserialize<Dictionary<string, object>>(jsonData);
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                throw new Exception(ex.Message);
            }
        }


        /// <summary>
        /// 二维数组转Dictionary<string, string>
        /// </summary>
        /// <param name="dArray"></param>
        /// <returns></returns>
        public static Dictionary<string, string> dArray2Dictionary(string[,] dArray)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                for (int i = 0; i < dArray.Length / 2; i++)
                {
                    if (dArray[i, 0].ToString().Equals(""))
                        return null;
                    dic.Add(dArray[i, 0].ToString(), dArray[i, 1].ToString());

                }
                return dic;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// Array转Dictionary<string, string>
        /// </summary>
        /// <param name="dArray"></param>
        /// <returns></returns>
        public static Dictionary<string, string> dArray2Dictionary(ArrayList dArray)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                for (int i = 0; i < dArray.Count; i++)
                {
                    object obj = dArray[i];
                    string[] sArr = (string[])obj;
                    if (sArr.Length != 2 || sArr[0].ToString().Equals(""))
                        return null;
                    dic.Add(sArr[0].ToString(), sArr[1].ToString());

                }
                return dic;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// object[]转Dictionary<string, string>
        /// </summary>
        /// <param name="dArray"></param>
        /// <returns></returns>
        public static Dictionary<string, string> dArray2Dictionary(object[] dArray)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                for (int i = 0; i < dArray.Length; i++)
                {
                    //WriteLog((dArray.Length).ToString(),"");
                    object[] obj = (object[])dArray[i];
                    if (obj.Length != 2)
                        return null;
                    string key = obj[0].ToString();
                    string value = obj[1].ToString();
                    //WriteLog(i+":["+key+": "+value+"]", "");
                    if (key.Equals(""))
                        return null;
                    dic.Add(key, value);

                }
                return dic;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// object[,]转Dictionary<string, string>
        /// </summary>
        /// <param name="dArray"></param>
        /// <returns></returns>
        public static Dictionary<string, string> dArray2Dictionary(object[,] dArray)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                for (int i = 0; i < dArray.Length / 2; i++)
                {
                    string key = dArray[i, 0].ToString();
                    string value = dArray[i, 1].ToString();
                    if (key.Equals(""))
                        return null;
                    dic.Add(key, value);

                }
                return dic;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// object[]转Dictionary<int, string[]>
        /// </summary>
        /// <param name="dArray"></param>
        /// <returns></returns>
        public static Dictionary<int, string[]> dArray2Dictionary2(object[] dArray)
        {
            Dictionary<int, string[]> dic = new Dictionary<int, string[]>();
            try
            {
                for (int i = 0; i < dArray.Length; i++)
                {
                    object[] obj = (object[])dArray[i];
                    string[] arrayTemp = new string[obj.Length];
                    for (int j = 0; j < obj.Length; j++)
                    {
                        arrayTemp[j] = obj[j].ToString();
                    }
                    dic.Add(i, arrayTemp);
                }
                return dic;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// object[,]转Dictionary<int, string[]>
        /// </summary>
        /// <param name="dArray">二维数组</param>
        /// <returns></returns>
        public static Dictionary<int, string[]> dArray2ToDictionary2(object[,] dArray)
        {
            Dictionary<int, string[]> dic = new Dictionary<int, string[]>();
            try
            {
                for (int i = 0; i < dArray.GetLength(0); i++)
                {
                    string[] arrayTemp = new string[dArray.GetLength(1)];
                    for (int j = 0; j < dArray.GetLength(1); j++)
                    {
                        arrayTemp[j] = dArray[i,j].ToString();
                    }
                    dic.Add(i, arrayTemp);
                }
                return dic;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// object[]转 object[,]
        /// </summary>
        /// <param name="dArray">object[]一维数组</param>
        /// <returns>object[,]二维数组</returns>
        public static object[,] dArray2Array2(object[] dArray)
        {
            try
            {
                object[,] array2 = new object[dArray.Length, ((object[])dArray[0]).Length];
                int length = dArray.Length;
                for (int i = 0; i < length; i++)
                {
                    object[] tempArray = (object[])dArray[i];
                    for (int j = 0; j < tempArray.Length; j++)
                    {
                        array2[i, j] = tempArray[j];
                    }
                }
                return array2;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }

        /// <summary>
        /// object[]转string[]
        /// </summary>
        /// <param name="dArray"></param>
        /// <returns></returns>
        public static string[] dArray2String1(object[] dArray)
        {
            List<string> list = new List<string>();
            string[] sl;
            try
            {
                for (int i = 0; i < dArray.Length; i++)
                {
                    list.Add(dArray[i].ToString());
                }
                sl = list.ToArray();
                return sl;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }
        #endregion

        #region NPOI DataGridView 导出 EXCEL
        /// <summary>
        /// NPOI DataGridView 导出 EXCEL
        /// </summary>
        /// <param name="fileName"> 默认保存文件名</param>
        /// <param name="dgv">DataGridView</param>
        /// <param name="fontname">字体名称</param>
        /// <param name="fontsize">字体大小</param>
        public void ExportExcel(string fileName, DataGridView dgv, string fontname, short fontsize)
        {
            //检测是否有数据
            if (dgv.Rows.Count == 0) return;
            //创建主要对象
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("Weight");
            //设置字体，大小，对齐方式
            HSSFCellStyle style = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFFont font = (HSSFFont)workbook.CreateFont();
            font.FontName = fontname;
            font.FontHeightInPoints = fontsize;

            font = setCellFontColor(font, "Red");//字体设为红色           
            style.SetFont(font);
            //style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            //style = setCellBGColor(style, "AUTOMATIC");//背景色为蓝色
            style.WrapText = true;
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//水平居中对齐
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//垂直居中


            //添加表头
            HSSFRow dataRow = (HSSFRow)sheet.CreateRow(0);
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                dataRow.CreateCell(i).SetCellValue(dgv.Columns[i].HeaderText);
                dataRow.GetCell(i).CellStyle = style;
            }
            //注释的这行是设置筛选的
            //sheet.SetAutoFilter(new CellRangeAddress(0, dgv.Columns.Count, 0, dgv.Columns.Count));
            //添加列及内容
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dataRow = (HSSFRow)sheet.CreateRow(i + 1);
                int[] maxColomnWidth = new int[dgv.Columns.Count];
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    maxColomnWidth[j] = sheet.GetColumnWidth(j);
                }
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    if (dgv.Rows[i].Cells[j].Value == null)
                    {
                        continue;
                    }
                    string ValueType = dgv.Rows[i].Cells[j].Value.GetType().ToString();
                    string Value = dgv.Rows[i].Cells[j].Value.ToString();
                    switch (ValueType)
                    {
                        case "System.String"://字符串类型
                            dataRow.CreateCell(j).SetCellValue(Value);
                            break;
                        case "System.DateTime"://日期类型
                            System.DateTime dateV;
                            System.DateTime.TryParse(Value, out dateV);
                            dataRow.CreateCell(j).SetCellValue(dateV);
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(Value, out boolV);
                            dataRow.CreateCell(j).SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(Value, out intV);
                            dataRow.CreateCell(j).SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(Value, out doubV);
                            dataRow.CreateCell(j).SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            dataRow.CreateCell(j).SetCellValue("");
                            break;
                        default:
                            dataRow.CreateCell(j).SetCellValue("");
                            break;
                    }
                    dataRow.GetCell(j).CellStyle = style;
                    //设置宽度
                    int nowColumnWidth = (Value.Length + 10) * 288;
                    if (nowColumnWidth > maxColomnWidth[j])
                    {
                        maxColomnWidth[j] = nowColumnWidth;
                        sheet.SetColumnWidth(j, maxColomnWidth[j]);
                    }
                }
            }

            //保存文件
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = fileName;
            MemoryStream ms = new MemoryStream();
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileName = saveDialog.FileName;
                if (!CheckFiles(saveFileName))
                {
                    workbook = null;
                    ms.Close();
                    ms.Dispose();
                    return;
                }
                workbook.Write(ms);
                FileStream file = new FileStream(saveFileName, FileMode.Create);
                workbook.Write(file);
                file.Close();
                workbook = null;
                ms.Close();
                ms.Dispose();
            }
            else
            {
                workbook = null;
                ms.Close();
                ms.Dispose();
            }
        }
        #endregion


        /// <summary>
        /// 为文件添加users，everyone用户组的完全控制权限
        /// </summary>
        /// <param name="filePath"></param>
        public static void AddSecurityControll2File(string filePath)
        {

            //获取文件信息
            FileInfo fileInfo = new FileInfo(filePath);
            //获得该文件的访问权限
            System.Security.AccessControl.FileSecurity fileSecurity = fileInfo.GetAccessControl();
            //添加ereryone用户组的访问权限规则 完全控制权限
            fileSecurity.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow));
            //添加Users用户组的访问权限规则 完全控制权限
            fileSecurity.AddAccessRule(new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow));
            //设置访问权限
            fileInfo.SetAccessControl(fileSecurity);
        }

        /// <summary>
        /// 保存excel文件,带路径选择对话框
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <param name="workbook">工作簿对象</param>
        public void saveExcelToNewFile(string filePath, IWorkbook workbook)
        {
            //保存文件
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = filePath;
            MemoryStream ms = new MemoryStream();
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileName = saveDialog.FileName;
                if (!CheckFiles(saveFileName))
                {
                    workbook = null;
                    ms.Close();
                    ms.Dispose();
                    return;
                }
                workbook.Write(ms);
                FileStream file = new FileStream(saveFileName, FileMode.Create);
                workbook.Write(file);
                file.Close();
                workbook = null;
                ms.Close();
                ms.Dispose();
            }
            else
            {
                ms.Close();
                ms.Dispose();
                return;
            }
            return;
        }

        ////保存excel文件,不打开文件选择框
        ///// <summary>
        ///// 保存excel文件,不打开文件选择框
        ///// </summary>
        ///// <param name="filePath">保存路径</param>
        ///// <param name="workbook">工作簿对象</param>
        //public void saveExcelWithoutAsk(string filePath, HSSFWorkbook workbook)
        //{
        //    if (!CheckFiles(filePath))
        //    {
        //        AddSecurityControll2File(filePath);
        //    }
        //    FileStream file = new FileStream(filePath, FileMode.Create,FileAccess.ReadWrite);
        //    workbook.Write(file);
        //    file.Close();
        //    file.Dispose();
        //    workbook = null;

        //    return;
        //}

        /// <summary>
        /// 保存excel文件,不打开文件选择框
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <param name="workbook">工作簿对象</param>
        public void saveExcelWithoutAsk(string filePath, IWorkbook workbook)
        {
            if (!CheckFiles(filePath))
            {
                AddSecurityControll2File(filePath);
            }
            FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            workbook.Write(file);
            file.Close();
            file.Dispose();
            workbook = null;

            return;
        }

        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="filePath">要删除的文件路径</param>
        public static void deleteFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            return;
        }


        #region 检测文件被占用
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);      

        /// <summary>
        /// 检测文件被占用 
        /// </summary>
        /// <param name="FileNames">要检测的文件路径</param>
        /// <returns></returns>
        public bool CheckFiles(string FileNames)
        {
            if (!File.Exists(FileNames))
            {
                //文件不存在
                return true;
            }
            IntPtr vHandle = _lopen(FileNames, OF_READWRITE | OF_SHARE_DENY_NONE);
            if (vHandle == HFILE_ERROR)
            {
                //文件被占用
                return false;
            }
            //文件没被占用
            CloseHandle(vHandle);
            return true;
        }
        #endregion


        /// <summary>
        /// 导出静态报表,数据为object[]格式
        /// </summary>
        /// <param name="modlePath">模板文件路径</param>
        /// <param name="sheetIndex">模板sheet索引</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="dArray">object[]数据</param>        
        /// <returns>是否成功</returns>
        public Boolean reportStaticExcel(string modlePath, int sheetIndex, string targetPath, object[] dArray)
        {
            string sheetName = loadExcelWorkbookI(modlePath).GetSheetName(sheetIndex);
            return reportStaticExcel(modlePath, sheetName, targetPath, dArray);
        }

        /// <summary>
        /// 导出静态报表,数据为object[]格式
        /// </summary>
        /// <param name="modlePath">模板文件路径</param>
        /// <param name="sheetName">模板sheet名</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="dArray">object[]数据</param>        
        /// <returns>是否成功</returns>
        public Boolean reportStaticExcel(string modlePath, string sheetName, string targetPath, object[] dArray)
        {
            try
            {
                Dictionary<string, string> hashMap = new Dictionary<string, string>();
                hashMap = dArray2Dictionary(dArray);
                fillDataToExcelByValue(modlePath, targetPath, hashMap, sheetName);
                return true;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return false;
            }
        }

        #region 废弃的写入图片方法
        ////导出带一组图片的excel
        ///// <summary>
        ///// 导出带一组图片的excel
        ///// </summary>
        ///// <param name="modlePath">模板文件路径</param>
        ///// <param name="targetPath">导出文件路径</param>
        ///// <param name="sheetIndex">sheet索引</param>
        ///// <param name="imgPathList">各个图片路径</param>
        ///// <param name="imgCellStr">图片位置标志字符串</param>
        ///// <returns>成功与否</returns>
        //public Boolean reportImagesExcel(string modlePath, string targetPath, int sheetIndex, object[] imgPathList, string imgCellStr )
        //{
        //    string sheetName = loadExcelWorkbook(modlePath).GetSheetName(sheetIndex);
        //    return reportImagesExcel( modlePath, targetPath, sheetName, imgPathList, imgCellStr );
        //}

        ////导出带一组图片的excel
        ///// <summary>
        ///// 导出带一组图片的excel
        ///// </summary>
        ///// <param name="modlePath">模板文件路径</param>
        ///// <param name="targetPath">导出文件路径</param>
        ///// <param name="sheetName">sheet名</param>
        ///// <param name="imgPathList">各个图片路径</param>
        ///// <param name="imgCellStr">图片位置标志字符串</param>
        ///// <returns>成功与否</returns>
        //public Boolean reportImagesExcel(string modlePath, string targetPath, string sheetName, object[] imgPathList, string imgCellStr )
        //{
        //    try
        //    {
        //        string[] sList = dArray2String1(imgPathList);
        //        HSSFWorkbook wb = loadExcelWorkbook(modlePath);
        //        HSSFSheet sheet = (HSSFSheet)wb.GetSheet(sheetName);
        //        wb = addImages2Excel(wb, sheetName, sList, imgCellStr);
        //        saveExcelWithoutAsk(targetPath, wb);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex,"");
        //        return false;
        //    }
        //}
        #endregion

        /// <summary>
        /// 查询值在sheet的位置
        /// </summary>
        /// <param name="sheet">工作表名</param>
        /// <param name="value">标志字符串</param>
        /// <returns>Point对象,X:行号,Y:列号</returns>
        public Point selectPosition(ISheet sheet, string value)
        {
            Point p = new Point();
            p.X = -1;
            p.Y = -1;
            try
            {
                int minRow = sheet.FirstRowNum;
                int maxRow = sheet.LastRowNum;
                for (int i = minRow; i <= maxRow; i++)
                {
                    IRow row = (IRow)sheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    for (int j = row.FirstCellNum; j <= row.LastCellNum; j++)
                    {
                        ICell cell = (ICell)row.GetCell(j);
                        if (cell == null)
                            continue;
                        string cellValue = getCellStringValueAllCase(cell);
                        if (cellValue.IndexOf(value) > -1)
                        {
                            p.X = i;
                            p.Y = j;
                            return p;
                        }
                    }
                }
                return p;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return p;
            }
        }

        /// <summary>
        /// 获取两个字符串中间的字符串
        /// </summary>
        /// <param name="s">目标字符串</param>
        /// <param name="s1">之前字符串</param>
        /// <param name="s2">之后字符串</param>
        /// <returns>两个字符串中间的字符串</returns>
        public static string Search_string(string s, string s1, string s2)
        {
            int n1, n2;
            n1 = s.IndexOf(s1, 0) + s1.Length;   //开始位置  
            n2 = s.IndexOf(s2, n1);               //结束位置  
            return s.Substring(n1, n2 - n1);   //取搜索的条数，用结束的位置-开始的位置,并返回  
        }

        /// <summary>
        /// 查询符合特征的标记字符串数组
        /// </summary>
        /// <param name="modlePath">模板路径</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="prefix">前缀</param>
        /// <param name="suffix">后缀</param>
        /// <returns>符合特征的标记字符串数组</returns>
        public string[] getModelMarks(string modlePath, int sheetIndex, string prefix, string suffix)
        {
            try
            {
                List<string> markList = new List<string>();
                ISheet sheet = loadExcelWorkbookI(modlePath).GetSheetAt(sheetIndex);

                int minRow = sheet.FirstRowNum;
                int maxRow = sheet.LastRowNum;
                for (int i = minRow; i <= maxRow; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null)
                        continue;
                    int minCol = row.FirstCellNum;
                    int maxCol = row.LastCellNum; ;
                    for (int j = minCol; j <= maxCol; j++)
                    {
                        ICell cell = sheet.GetRow(i).GetCell(j);
                        if (cell == null)
                            continue;
                        string cellValue = getCellStringValueAllCase(cell);
                        //可能是一个单元格含有多个标记字段,需要做循环
                        while (cellValue.IndexOf(prefix) > -1 && cellValue.IndexOf(suffix) > -1)
                        {
                            markList.Add(Search_string(cellValue, prefix, suffix));
                            cellValue = cellValue.Remove(cellValue.IndexOf(prefix), 1);
                            cellValue = cellValue.Remove(cellValue.IndexOf(suffix), 1);
                        }
                    }
                }
                return markList.ToArray();
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return null;
            }
        }
                
        /// <summary>
        /// 获取字符串数组在sheet中的列号
        /// </summary>
        /// <param name="sheet">目标sheet</param>
        /// <param name="value">目标字符串数组</param>
        /// <param name="prefix">前缀</param>
        /// <param name="suffix">后缀</param>
        /// <returns>字符串数组在sheet中的列号</returns>
        public int[] getArraySequen(ISheet sheet, string[] value, string prefix, string suffix)
        {
            int[] sequen = new int[value.Length];
            for (int i = 0; i < value.Length; i++)
            {
                sequen[i] = selectPosition(sheet, prefix + value[i] + suffix).Y;
            }
            return sequen;
        }

        /// <summary>
        /// 按照顺序数组排列二维数组的列
        /// </summary>
        /// <param name="sequen">目标顺序</param>
        /// <param name="array">目标二维数组</param>
        /// <returns>重排列顺序的二维数组</returns>
        public static object[,] getSequenArray2(int[] sequen, object[,] array)
        {
            try
            {
                int rowCount = array.GetLength(0);
                int colCount = array.GetLength(1);
                object[,] sequenArray2 = new object[rowCount, colCount];
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        sequenArray2[i, sequen[j]] = array[i, j];
                    }
                }
                return sequenArray2;
            }
            catch (Exception ex)
            {
                WriteLog(ex,"");
                return array;
            }
        }


        /// <summary>
        /// 设置替换单元格的值
        /// </summary>
        /// <param name="sheet">源工作表</param>
        /// <param name="sourceString">标志字符串</param>
        /// <param name="targetString">要替换的值</param>
        /// <returns>目标工作表</returns>
        public ISheet replaceCellValue(ISheet sheet, string sourceString, string targetString)
        {
            Point p = selectPosition(sheet, targetString);
            //-1代表没找到标志字符串
            if (p.X == -1 || p.Y == -1)
            {
                return sheet;
            }
            ICell cell = (ICell)sheet.GetRow(p.X).GetCell(p.Y);
            string cellValue = getCellStringValueAllCase(cell);
            cellValue = cellValue.Replace(targetString, sourceString);
            cell.SetCellValue(cellValue);
            return sheet;
        }

        /// <summary>
        /// 设置替换单元格的值
        /// </summary>
        /// <param name="workbook">源工作簿</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="sourceString">标志字符串</param>
        /// <param name="targetString">要替换的值</param>
        /// <returns>目标工作簿</returns>
        public IWorkbook replaceCellValue(IWorkbook workbook, int sheetIndex, string sourceString, string targetString)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            Point p = selectPosition(sheet, targetString);
            //-1代表没找到标志字符串
            if (p.X == -1 || p.Y == -1)
            {
                return workbook;
            }
            ICell cell = (ICell)sheet.GetRow(p.X).GetCell(p.Y);
            string cellValue = getCellStringValueAllCase(cell);
            cellValue = cellValue.Replace(targetString, sourceString);
            cell.SetCellValue(cellValue);
            return workbook;
        }

        /// <summary>
        /// 设置替换单元格的值
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="newString">标志字符串</param>
        /// <param name="oldString">要替换的值</param>
        public void replaceLastCellValue(string filePath, string newString, string oldString)
        {
            IWorkbook workbook = loadExcelWorkbookI(filePath);
            replaceLastCellValue(workbook, newString, oldString);
            saveExcelWithoutAsk(filePath, workbook);
        }

        /// <summary>
        /// 设置替换单元格的值
        /// </summary>
        /// <param name="workbook">源工作簿</param>
        /// <param name="newString">标志字符串</param>
        /// <param name="oldString">要替换的值</param>
        /// <returns>目标工作簿</returns>
        public IWorkbook replaceLastCellValue(IWorkbook workbook, string newString, string oldString)
        {
            //从后向前查找,替换第一个遇到的flag
            for(int i= workbook.NumberOfSheets-1; i >= 0; i--)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                Point p = selectPosition(sheet, oldString);
                //-1代表没找到标志字符串
                if (p.X == -1 || p.Y == -1)
                {
                    continue;
                }
                else
                {
                    ICell cell = (ICell)sheet.GetRow(p.X).GetCell(p.Y);
                    string cellValue = getCellStringValueAllCase(cell);
                    cellValue = cellValue.Replace(oldString, newString);
                    cell.SetCellValue(cellValue);
                    break;
                }            
            }
            return workbook;
        }

        #region 废弃不用的导出方法

        ////导出一维报表从表,数据为object[]格式
        ///// <summary>
        ///// 导出一维报表(不扩展列,只扩展行)从表,数据为object[]格式
        ///// </summary>
        ///// <param name="modlePath">模板路径</param>
        ///// <param name="sheetIndex">模板模板sheet索引名</param>
        ///// <param name="cellPosiValue">起始单元格标记字符串</param>
        ///// <param name="targetPath">报表文件路径</param>
        ///// <param name="dArray">二维数据</param>
        ///// <param name="colListC">需要合并的列</param>
        ///// <param name="projCol">项目名所在的列号</param>
        ///// <returns>是否成功</returns>
        //public Boolean reportOneDimDExcel(string modlePath, int sheetIndex, string cellPosiValue, string targetPath, object[] dArray, object[] colListC,  int projCol )
        //{
        //    string sheetName = loadExcelWorkbook(modlePath).GetSheetName(sheetIndex);
        //    return reportOneDimDExcel(modlePath,sheetName,cellPosiValue,targetPath,dArray,colListC,projCol);
        //}

        ////导出一维报表从表,数据为object[]格式
        ///// <summary>
        ///// 导出一维报表(不扩展列,只扩展行)从表,数据为object[]格式
        ///// </summary>
        ///// <param name="modlePath">模板路径</param>
        ///// <param name="sheetName">模板sheet名</param>
        ///// <param name="cellPosiValue">起始单元格标记字符串</param>
        ///// <param name="targetPath">报表文件路径</param>
        ///// <param name="dArray">二维数据</param>
        ///// <param name="colListC">需要合并的列</param>
        ///// <param name="projCol">项目名所在的列号</param>
        ///// <returns>是否成功</returns>
        //public Boolean reportOneDimDExcel(string modlePath, string sheetName, string cellPosiValue, string targetPath, object[] dArray, object[] colListC, int projCol )
        //{
        //    try
        //    {
        //        int[] colList = new int[colListC.Length];
        //        for (int i = 0; i < colListC.Length; i++)
        //        {
        //            colList[i] = int.Parse(colListC[i].ToString());
        //        }
        //        HSSFWorkbook wb = loadExcelWorkbook(modlePath);//获取workbook
        //        HSSFSheet sheet = (HSSFSheet)wb.GetSheet(sheetName);//获取sheet
        //        Dictionary<int, string[]> dic = dArray2Dictionary2(dArray);

        //        int row = selectPosition(sheet, cellPosiValue).X;//起始单元格行号
        //        int col = selectPosition(sheet, cellPosiValue).Y; //起始单元格列号
        //        for (int i = row; i < row + dic.Count; i++)
        //        {

        //            #region 扩充行,并设置格式为上一行
        //            //先扩充一行
        //            sheet.ShiftRows(i + 1,                                 //--开始行
        //                sheet.LastRowNum,                            //--结束行
        //                1,                             //--移动大小(行数)--往下移动
        //                true,                                   //是否复制行高
        //                false,                                  //是否重置行高
        //                true                                    //是否移动批注
        //                );
        //            // 对批量移动后空出的空行插，创建相应的行，并以插入行的上一行为格式源(即：InsertRowIndex-1的那一行)

        //            HSSFRow targetRow = null;
        //            HSSFCell sourceCell = null;
        //            HSSFCell targetCell = null;
        //            HSSFRow mySourceStyleRow = (HSSFRow)sheet.GetRow(i);
        //            if (mySourceStyleRow == null)
        //                continue;

        //            targetRow = (HSSFRow)sheet.CreateRow(i + 1);

        //            for (int m = mySourceStyleRow.FirstCellNum; m < mySourceStyleRow.LastCellNum; m++)
        //            {
        //                sourceCell = (HSSFCell)mySourceStyleRow.GetCell(m);
        //                if (sourceCell == null)
        //                    continue;
        //                targetCell = (HSSFCell)targetRow.CreateCell(m);
        //                //targetCell.Encoding = sourceCell.Encoding;
        //                targetCell.CellStyle = sourceCell.CellStyle;
        //                targetCell.SetCellType(sourceCell.CellType);

        //            }

        //            #endregion

        //            string[] arrayTemp = dic[i - row];
        //            for (int j = col; j < col + arrayTemp.Length; j++)
        //            {
        //                HSSFCell cell = (HSSFCell)sheet.GetRow(i + 1).GetCell(j);
        //                if (cell == null)
        //                    continue;
        //                cell.SetCellValue(arrayTemp[j - col]);
        //            }
        //        }

        //        //删除起始行,整体往上移动1行
        //        sheet.ShiftRows(row + 1, sheet.LastRowNum, -1);

        //        //计算需要合并的区域,执行合并单元格
        //        mergeCells(wb, sheetName, colList, projCol, row, row + dic.Count, targetPath);

        //        saveExcelWithoutAsk(targetPath, wb);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return false;
        //    }
        //}

        
            
        ///// <summary>
        ///// 导出一维报表(不扩展列,只扩展行)从表,数据为object[]格式
        ///// </summary>
        ///// <param name="modlePath">模板文件路径</param>
        ///// <param name="sheetIndex"></param>
        ///// <param name="targetPath">目标文件路径</param>
        ///// <param name="dArray">二维数组,第一行为表头</param>
        ///// <param name="colListC"></param>
        ///// <returns></returns>
        //public Boolean reportOneDimDExcel(string modlePath, int sheetIndex, string targetPath, object[] dArray, object[] colListC)
        //{
        //    try
        //    {             
        //        HSSFWorkbook wb = loadExcelWorkbook(modlePath);//获取workbook
        //        HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(sheetIndex);//获取sheet
        //        int[] colseq = getArraySequen(sheet, dArray2String1(colListC), "&[", "]"); //获取合并列列名在模板中的顺序
        //        string[] tableHead = dArray2String1((object[])dArray[0]);//最开始的行为表头
        //        int[] colHeadSeq = getArraySequen(sheet, tableHead, "&[", "]"); //获取表头在模板中的顺序,表头数据无标记符号&[],需要添加
        //        object[,] seqArray2 = getSequenArray2(colHeadSeq, dArray2Array2(dArray));//获取排序后的二维数组
        //        Dictionary<int, string[]> dic = dArray2ToDictionary2(seqArray2);

        //        ArrayList arr = new ArrayList(colHeadSeq);    //声明一个ArrayList并载入数组
        //        int index = arr.IndexOf(0);          //通过indexof函数找到0所在数组中的位置,此处即是表的起始位置
        //        string cellPosiValue = tableHead[index];

        //        int row = selectPosition(sheet, cellPosiValue).X;//起始单元格行号
        //        int col = selectPosition(sheet, cellPosiValue).Y; //起始单元格列号

        //        for (int i = row; i < row + dic.Count; i++)
        //        {

        //            #region 扩充行,并设置格式为上一行
        //            //先扩充一行
        //            sheet.ShiftRows(i + 1,                                 //--开始行
        //                sheet.LastRowNum,                            //--结束行
        //                1,                             //--移动大小(行数)--往下移动
        //                true,                                   //是否复制行高
        //                false,                                  //是否重置行高
        //                true                                    //是否移动批注
        //                );
        //            // 对批量移动后空出的空行插，创建相应的行，并以插入行的上一行为格式源(即：InsertRowIndex-1的那一行)

        //            HSSFRow targetRow = null;
        //            HSSFCell sourceCell = null;
        //            HSSFCell targetCell = null;
        //            HSSFRow mySourceStyleRow = (HSSFRow)sheet.GetRow(i);
        //            if (mySourceStyleRow == null)
        //                continue;

        //            targetRow = (HSSFRow)sheet.CreateRow(i + 1);

        //            for (int m = mySourceStyleRow.FirstCellNum; m < mySourceStyleRow.LastCellNum; m++)
        //            {
        //                sourceCell = (HSSFCell)mySourceStyleRow.GetCell(m);
        //                if (sourceCell == null)
        //                    continue;
        //                targetCell = (HSSFCell)targetRow.CreateCell(m);
        //                //targetCell.Encoding = sourceCell.Encoding;
        //                targetCell.CellStyle = sourceCell.CellStyle;
        //                targetCell.SetCellType(sourceCell.CellType);

        //            }

        //            #endregion

        //            string[] arrayTemp = dic[i - row];
        //            for (int j = col; j < col + arrayTemp.Length; j++)
        //            {
        //                HSSFCell cell = (HSSFCell)sheet.GetRow(i + 1).GetCell(j);
        //                if (cell == null)
        //                    continue;
        //                cell.SetCellValue(arrayTemp[j - col]);
        //            }
        //        }

        //        //删除起始行,整体往上移动1行
        //        sheet.ShiftRows(row + 1, sheet.LastRowNum, -1);                

        //        //计算需要合并的区域,执行按行合并单元格
        //        sheet = mergeRowCells(sheet);
        //        //计算需要合并的区域,执行按列合并单元格
        //        sheet = mergeCells(sheet, colseq, row, row + dic.Count);


        //        saveExcelWithoutAsk(targetPath, wb);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return false;
        //    }
        //}
        #endregion

        /// <summary>
        /// 导出一维报表(不扩展列,只扩展行)从表,数据为object[]格式
        /// </summary>
        /// <param name="modlePath">模板文件路径</param>
        /// <param name="sheetIndex"></param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="dArray">二维数组,第一行为表头</param>
        /// <param name="colListC"></param>
        /// <returns></returns>
        public Boolean reportOneDimDExcel(string modlePath, int sheetIndex, string targetPath, object[] dArray, object[] colListC)
        {
            try
            {
                IWorkbook wb = loadExcelWorkbookI(modlePath);//获取workbook
                ISheet sheet = wb.GetSheetAt(sheetIndex);//获取sheet
                int[] colseq = getArraySequen(sheet, dArray2String1(colListC), "&[", "]"); //获取合并列列名在模板中的顺序
                string[] tableHead = dArray2String1((object[])dArray[0]);//最开始的行为表头
                int[] colHeadSeq = getArraySequen(sheet, tableHead, "&[", "]"); //获取表头在模板中的顺序,表头数据无标记符号&[],需要添加
                object[,] seqArray2 = getSequenArray2(colHeadSeq, dArray2Array2(dArray));//获取排序后的二维数组
                Dictionary<int, string[]> dic = dArray2ToDictionary2(seqArray2);

                ArrayList arr = new ArrayList(colHeadSeq);    //声明一个ArrayList并载入数组
                int index = arr.IndexOf(0);          //通过indexof函数找到0所在数组中的位置,此处即是表的起始位置
                string cellPosiValue = tableHead[index];

                int row = selectPosition(sheet, cellPosiValue).X;//起始单元格行号
                int col = selectPosition(sheet, cellPosiValue).Y; //起始单元格列号

                for (int i = row; i < row + dic.Count; i++)
                {

                    #region 扩充行,并设置格式为上一行
                    //先扩充一行
                    sheet.ShiftRows(i + 1,                                 //--开始行
                        sheet.LastRowNum,                            //--结束行
                        1,                             //--移动大小(行数)--往下移动
                        true,                                   //是否复制行高
                        false                                  //是否重置行高
                        );
                    // 对批量移动后空出的空行插，创建相应的行，并以插入行的上一行为格式源(即：InsertRowIndex-1的那一行)

                    IRow targetRow = null;
                    ICell sourceCell = null;
                    ICell targetCell = null;
                    IRow mySourceStyleRow = sheet.GetRow(i);
                    if (mySourceStyleRow == null)
                        continue;

                    targetRow = sheet.CreateRow(i + 1);

                    for (int m = mySourceStyleRow.FirstCellNum; m < mySourceStyleRow.LastCellNum; m++)
                    {
                        sourceCell = mySourceStyleRow.GetCell(m);
                        if (sourceCell == null)
                            continue;
                        targetCell = targetRow.CreateCell(m);
                        targetCell.CellStyle = sourceCell.CellStyle;
                        targetCell.SetCellType(sourceCell.CellType);

                    }

                    #endregion

                    string[] arrayTemp = dic[i - row];
                    for (int j = col; j < col + arrayTemp.Length; j++)
                    {
                        ICell cell = sheet.GetRow(i + 1).GetCell(j);
                        if (cell == null)
                            continue;
                        cell.SetCellValue(arrayTemp[j - col]);
                    }
                }

                //删除起始行,整体往上移动1行
                sheet.ShiftRows(row + 1, sheet.LastRowNum, -1);

                //计算需要合并的区域,执行按行合并单元格
                sheet = mergeRowCells(sheet);
                //计算需要合并的区域,执行按列合并单元格
                sheet = mergeCells(sheet, colseq, row, row + dic.Count);


                saveExcelWithoutAsk(targetPath, wb);

                return true;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return false;
            }
        }
        
        /// <summary>
        /// 合并指定列,按值相等合并
        /// </summary>
        /// <param name="sheetName">目标sheet</param>
        /// <param name="colList">要合并的单元格所在列</param>
        /// <param name="startRow">开始行</param>
        /// <param name="endRow">结束行</param>
        public ISheet mergeCells(ISheet sheet, int[] colList, int startRow, int endRow)
        {
            try
            {
                for (int i = 0; i < colList.Length; i++)//遍历需要合并的列
                {
                    #region 检查数值相等并合并当前列
                    IRow tempHSSFRow = sheet.GetRow(startRow);
                    if (tempHSSFRow == null)
                        continue;
                    ICell tempHSSFCell = tempHSSFRow.GetCell(colList[i]);
                    if (tempHSSFCell == null)
                        continue;
                    string tempCellValue = getCellStringValueAllCase(tempHSSFCell);//最新单元格值
                    string tempTESTNO = getCellStringValueAllCase(tempHSSFRow.GetCell(2));//最新检测项值
                    int tempRow = startRow;//最新行号,作为需要合并的起始行
                    int beforeRow = startRow;//之前的行号,作为需要合并的结束行
                    for (int j = startRow + 1; j <= endRow; j++)//遍历列的指定行集合
                    {
                        IRow tempHSSFRow1 = sheet.GetRow(j);
                        if (tempHSSFRow1 == null)
                            continue;
                        ICell tempHSSFCell1 = tempHSSFRow1.GetCell(colList[i]);
                        if (tempHSSFCell1 == null)
                            continue;
                        string nowCellValue = getCellStringValueAllCase(tempHSSFCell1);//目前单元格值
                        string nowTESTNO = getCellStringValueAllCase(tempHSSFRow1.GetCell(2));//目前检测项值

                        //如果相等则之前的行号+1,需要根据检测项相等判定合并,检测项默认在第3列
                        if (tempCellValue.Equals(nowCellValue) && tempTESTNO.Equals(nowTESTNO))
                        {
                            beforeRow++;

                        }
                        else//如果不等则合并记录下的单元格区域,并记录新的行号和单元格值
                        {
                            if (tempRow < beforeRow)//如果最新行号小于遍历的上一个行号
                            {

                                //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
                                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(tempRow, beforeRow, colList[i], colList[i]));

                            }
                            tempCellValue = getCellStringValueAllCase(sheet.GetRow(j).GetCell(colList[i]));//更新单元格值
                            tempTESTNO = getCellStringValueAllCase(sheet.GetRow(j).GetCell(2));//更新检测项值
                            tempRow = j;
                            beforeRow++;
                        }


                        #endregion
                    }
                }
                return sheet;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return sheet;
            }
        }

        //通过坐标获取单元格所在合并区域
        /// <summary>
        /// 通过坐标获取单元格所在合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">行号</param>
        /// <param name="col">列号</param>
        /// <returns>2x2数组</returns>
        public int[,] getCellMergeArea(ISheet sheet, int row, int col)
        {
            int[,] range = new int[2, 2] { { -1, -1 }, { -1, -1 } };
            ICell cell = sheet.GetRow(row).GetCell(col);

            //如果不是被合并的单元格,返回当前位置区域
            if (!cell.IsMergedCell)
                return new int[2, 2] { { row, col }, { row, col } };

            ////如果是被合并的单元格
            //NPOI.SS.Util.CellRangeAddress cra = cell.ArrayFormulaRange;
            //range[0, 0] = cra.FirstRow;
            //range[0, 1] = cra.FirstColumn;
            //range[1, 0] = cra.LastRow;
            //range[1, 1] = cra.LastColumn;
            ////return range;

            for (int i = 0; i < sheet.NumMergedRegions; i++) // 循环所有合并的单元格
            {
                var mergeArea = sheet.GetMergedRegion(i);
                if (mergeArea.FirstRow == row && mergeArea.FirstColumn == col)
                {
                    range[0, 0] = mergeArea.FirstRow;
                    range[0, 1] = mergeArea.FirstColumn;
                    range[1, 0] = mergeArea.LastRow;
                    range[1, 1] = mergeArea.LastColumn;
                    break;
                }

            }

            return range;
            


        }

        /// <summary>
        /// 通过标记字符串获取单元格所在合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">行号</param>
        /// <param name="col">列号</param>
        /// <returns>2x2数组</returns>
        public int[,] getCellMergeArea(ISheet sheet, string flag)
        {
            int row = selectPosition(sheet, flag).X;
            int col = selectPosition(sheet, flag).Y;
            //ICell cell = sheet.GetRow(row).GetCell(col);
            return getCellMergeArea(sheet, row, col);



        }

        

        /// <summary>
        /// 合并指定行,按值相等合并
        /// </summary>
        /// <param name="sheet">目标sheet</param>
        /// <param name="startRow">开始行</param>
        /// <param name="endRow">结束行</param>
        /// <param name="startCol">来时列</param>
        /// <param name="endCol">结束列</param>
        public ISheet mergeRowCells(ISheet sheet)
        {
            try
            {
                int startRow = sheet.FirstRowNum;
                int endRow = sheet.LastRowNum; ;
                for (int i = startRow; i < endRow; i++)//遍历需要合并的行
                {
                    #region 检查数值相等并合并当前列
                    IRow tempHSSFRow = sheet.GetRow(i);
                    if (tempHSSFRow == null)
                        continue;
                    ICell tempHSSFCell = tempHSSFRow.GetCell(Int32.Parse(tempHSSFRow.FirstCellNum.ToString()));
                    if (tempHSSFCell == null)
                        continue;
                    string tempCellValue = getCellStringValueAllCase(tempHSSFCell);//最新单元格值
                    if (tempCellValue == null)
                    {
                        tempCellValue = "";
                    }
                    int tempCol = tempHSSFRow.FirstCellNum;//最新列号,作为需要合并的起始列
                    int beforeCol = tempHSSFRow.FirstCellNum;//之前的列号,作为需要合并的结束列
                    for (int j = tempHSSFRow.FirstCellNum + 1; j <= tempHSSFRow.LastCellNum; j++)//遍历列的指定行集合
                    {

                        IRow tempHSSFRow1 = sheet.GetRow(i);
                        if (tempHSSFRow1 == null)
                            continue;
                        ICell tempHSSFCell1 = tempHSSFRow1.GetCell(j);
                        if (tempHSSFCell1 == null)
                            continue;
                        string nowCellValue = getCellStringValueAllCase(tempHSSFCell1);//目前单元格值
                        if (nowCellValue == null)
                        {
                            nowCellValue = "";
                        }

                        if (tempCellValue.Equals(nowCellValue))//如果相等则之前的列号+1
                        {
                            beforeCol++;

                        }
                        else//如果不等则合并记录下的单元格区域,并记录新的列号和单元格值
                        {
                            //如果最新列号小于遍历的上一个列号,且单元格值非空
                            if (tempCol < beforeCol && !tempCellValue.Equals(""))
                            {

                                //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
                                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, tempCol, beforeCol));

                            }
                            tempCellValue = getCellStringValueAllCase((ICell)sheet.GetRow(i).GetCell(j));//更新单元格值
                            tempCol = j;
                            beforeCol++;
                        }


                        #endregion
                    }
                }
                return sheet;
            }
            catch (Exception ex)
            {
                WriteLog(ex, "");
                return sheet;
            }
        }

        /// <summary>
        /// 获取单元格的String类型值
        /// </summary>
        /// <param name="tCell">目标单元格</param>
        /// <returns>单元格的String类型值</returns>
        public static string getCellStringValueAllCase(ICell tCell)
        {
            string tempValue = "";
            switch (tCell.CellType)
            {
                case NPOI.SS.UserModel.CellType.Blank:
                    break;
                case NPOI.SS.UserModel.CellType.Boolean:
                    tempValue = tCell.BooleanCellValue.ToString();
                    break;
                case NPOI.SS.UserModel.CellType.Error:
                    break;
                case NPOI.SS.UserModel.CellType.Formula:
                    NPOI.SS.UserModel.IFormulaEvaluator fe = NPOI.SS.UserModel.WorkbookFactory.CreateFormulaEvaluator(tCell.Sheet.Workbook);
                    var cellValue = fe.Evaluate(tCell);
                    switch (cellValue.CellType)
                    {
                        case NPOI.SS.UserModel.CellType.Blank:
                            break;
                        case NPOI.SS.UserModel.CellType.Boolean:
                            tempValue = cellValue.BooleanValue.ToString();
                            break;
                        case NPOI.SS.UserModel.CellType.Error:
                            break;
                        case NPOI.SS.UserModel.CellType.Formula:
                            break;
                        case NPOI.SS.UserModel.CellType.Numeric:
                            tempValue = cellValue.NumberValue.ToString();
                            break;
                        case NPOI.SS.UserModel.CellType.String:
                            tempValue = cellValue.StringValue.ToString();
                            break;
                        case NPOI.SS.UserModel.CellType.Unknown:
                            break;
                        default:
                            break;
                    }
                    break;
                case NPOI.SS.UserModel.CellType.Numeric:

                    if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(tCell))
                    {
                        tempValue = tCell.DateCellValue.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        tempValue = tCell.NumericCellValue.ToString();
                    }
                    break;
                case NPOI.SS.UserModel.CellType.String:
                    tempValue = tCell.StringCellValue.Trim();
                    break;
                case NPOI.SS.UserModel.CellType.Unknown:
                    break;
                default:
                    break;
            }
            return tempValue;
        }

        ////获取单元格的String类型值
        ///// <summary>
        ///// 获取单元格的String类型值
        ///// </summary>
        ///// <param name="tCell">目标单元格</param>
        ///// <returns>单元格的String类型值</returns>
        //public static string getCellStringValueAllCase(HSSFCell tCell)
        //{
        //    string tempValue = "";
        //    switch (tCell.CellType)
        //    {
        //        case NPOI.SS.UserModel.CellType.Blank:
        //            break;
        //        case NPOI.SS.UserModel.CellType.Boolean:
        //            tempValue = tCell.BooleanCellValue.ToString();
        //            break;
        //        case NPOI.SS.UserModel.CellType.Error:
        //            break;
        //        case NPOI.SS.UserModel.CellType.Formula:
        //            NPOI.SS.UserModel.IFormulaEvaluator fe = NPOI.SS.UserModel.WorkbookFactory.CreateFormulaEvaluator(tCell.Sheet.Workbook);
        //            var cellValue = fe.Evaluate(tCell);
        //            switch (cellValue.CellType)
        //            {
        //                case NPOI.SS.UserModel.CellType.Blank:
        //                    break;
        //                case NPOI.SS.UserModel.CellType.Boolean:
        //                    tempValue = cellValue.BooleanValue.ToString();
        //                    break;
        //                case NPOI.SS.UserModel.CellType.Error:
        //                    break;
        //                case NPOI.SS.UserModel.CellType.Formula:
        //                    break;
        //                case NPOI.SS.UserModel.CellType.Numeric:
        //                    tempValue = cellValue.NumberValue.ToString();
        //                    break;
        //                case NPOI.SS.UserModel.CellType.String:
        //                    tempValue = cellValue.StringValue.ToString();
        //                    break;
        //                case NPOI.SS.UserModel.CellType.Unknown:
        //                    break;
        //                default:
        //                    break;
        //            }
        //            break;
        //        case NPOI.SS.UserModel.CellType.Numeric:

        //            if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(tCell))
        //            {
        //                tempValue = tCell.DateCellValue.ToString("yyyy-MM-dd");
        //            }
        //            else
        //            {
        //                tempValue = tCell.NumericCellValue.ToString();
        //            }
        //            break;
        //        case NPOI.SS.UserModel.CellType.String:
        //            tempValue = tCell.StringCellValue.Trim();
        //            break;
        //        case NPOI.SS.UserModel.CellType.Unknown:
        //            break;
        //        default:
        //            break;
        //    }
        //    return tempValue;
        //}



        public void alert(object msg)
        {
            MessageBox.Show(msg.ToString());
        }

        //废弃的使用HSSFSheet的方法,改用ISheet后可兼容xls和xlsx
        #region 废弃的使用HSSFSheet的方法,改用ISheet后可兼容xls和xlsx
        ///// <summary>
        ///// 设置打印标题区间
        ///// </summary>
        ///// <param name="filePath"></param>
        ///// <param name="sheetIndex"></param>
        ///// <param name="startRowFlag">起始行所在标记,注意此行不计入表头</param>
        ///// <param name="endRowflag">结束行标记</param>
        ///// <param name="startCol">起始列</param>
        ///// <param name="endCol">结束列</param>
        ///// <returns>2个元素的数组,表头行区间</returns>
        //public int[] SetTableHeader(string filePath, int sheetIndex, string startRowFlag, string endRowflag, int startCol, int endCol)
        //{
        //    HSSFWorkbook wb = loadExcelWorkbook(filePath);
        //    HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(sheetIndex);

        //    //数据写完后会多一个空行,需要手动删掉
        //    sheet.RemoveRow(sheet.GetRow(sheet.LastRowNum));

        //    int startRow = selectPosition(sheet, startRowFlag).X+1;
        //    int endRow = selectPosition(sheet, endRowflag).X;
        //    //设置打印标题用,CellRangeAddress参数:(起始行号，终止行号， 起始列号，终止列号)
        //    sheet.RepeatingRows = new NPOI.SS.Util.CellRangeAddress(startRow, endRow, startCol, endCol);
        //    saveExcelWithoutAsk(filePath,wb);
        //    int[] range = new int[]{startRow,endRow};
        //    return range;
        //}

        ///// <summary>
        ///// 读取指定xls文件的指定sheet
        ///// </summary>
        ///// <param name="filePath">文件路径</param>
        ///// <param name="sheeIndex">工作表sheet索引名</param>
        ///// <returns>读取到的工作表,失败返回空</returns>
        //public HSSFSheet loadExcelSheet(string filePath, int sheeIndex)
        //{
        //    try
        //    {
        //        FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        //        HSSFWorkbook wb = new HSSFWorkbook(file);
        //        HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(sheeIndex);
        //        file.Close();
        //        return sheet;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return null;
        //    }
        //}

        ////创建新的sheet
        ///// <summary>
        ///// 创建新的sheet,如果目标索引存在,则此sheet重命名为目标sheet名,不存在则创建
        ///// </summary>
        ///// <param name="workbook">工作簿</param>
        ///// <param name="sheetIndex">目标sheet索引</param>
        ///// <param name="sheetName">sheet名</param>
        ///// <param name="fromSheet">源工作表</param>
        //public void addNewSheet(HSSFWorkbook workbook, int sheetIndex, string sheetName, HSSFSheet fromSheet)
        //{
        //    HSSFSheet toSheet;
        //    try
        //    {
        //        try 
        //        {
        //            toSheet = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
        //            workbook.SetSheetName(sheetIndex, sheetName);
        //        }
        //        catch 
        //        {
        //            workbook.CreateSheet(sheetName);
        //            toSheet = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
        //        }
        //        return;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return;
        //    }

        //}


        ////查询值在sheet的位置
        ///// <summary>
        ///// 查询值在sheet的位置
        ///// </summary>
        ///// <param name="sheet">工作表名</param>
        ///// <param name="value">标志字符串</param>
        ///// <returns>Point对象,X:行号,Y:列号</returns>
        //public Point selectPosition(HSSFSheet sheet, string value)
        //{
        //    Point p = new Point();
        //    p.X = -1;
        //    p.Y = -1;
        //    try
        //    {               
        //        int minRow = sheet.FirstRowNum;
        //        int maxRow = sheet.LastRowNum;
        //        for (int i = minRow; i <= maxRow; i++)
        //        {
        //            HSSFRow row = (HSSFRow)sheet.GetRow(i);
        //            if (row == null)
        //            {
        //                continue;
        //            }
        //            for (int j = row.FirstCellNum; j <= row.LastCellNum; j++)
        //            {
        //                HSSFCell cell = (HSSFCell)row.GetCell(j);
        //                if (cell == null)
        //                    continue;
        //                string cellValue = getCellStringValueAllCase(cell);
        //                if (cellValue.IndexOf(value) > -1)
        //                {
        //                    p.X = i;
        //                    p.Y = j;
        //                    return p;
        //                }
        //            }
        //        }
        //        return p;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return p;
        //    }
        //}
        ///// <summary>
        ///// 获取字符串数组在sheet中的列号
        ///// </summary>
        ///// <param name="sheet">目标sheet</param>
        ///// <param name="value">目标字符串数组</param>
        ///// <param name="prefix">前缀</param>
        ///// <param name="suffix">后缀</param>
        ///// <returns>字符串数组在sheet中的列号</returns>
        //public int[] getArraySequen(HSSFSheet sheet, string[] value, string prefix, string suffix){
        //    int[] sequen = new int[value.Length];
        //    for (int i = 0; i < value.Length; i++)
        //    {
        //        sequen[i] = selectPosition(sheet, prefix + value[i] + suffix).Y;
        //    }
        //    return sequen;
        //}

        ///// <summary>
        ///// 设置替换单元格的值
        ///// </summary>
        ///// <param name="sheet">源工作表</param>
        ///// <param name="sourceString">标志字符串</param>
        ///// <param name="targetString">要替换的值</param>
        ///// <returns>目标工作表</returns>
        //public HSSFSheet replaceCellValue(HSSFSheet sheet, string sourceString, string targetString)
        //{
        //    Point p = selectPosition(sheet, targetString);
        //    //-1代表没找到标志字符串
        //    if (p.X == -1 || p.Y == -1)
        //    {
        //        return sheet;
        //    }
        //    HSSFCell cell = (HSSFCell)sheet.GetRow(p.X).GetCell(p.Y);
        //    string cellValue = getCellStringValueAllCase(cell);
        //    cellValue = cellValue.Replace(targetString, sourceString);
        //    cell.SetCellValue(cellValue);
        //    return sheet;
        //}


        ////合并指定列,按项目名相等
        ////需求是只要项目名相等,则其他需要合并的列数据也一定相等
        ///// <summary>
        ///// 合并指定列,按值相等合并
        ///// </summary>
        ///// <param name="wb">目标workbook</param>
        ///// <param name="sheetIndex">目标sheet索引</param>
        ///// <param name="colList">要合并的单元格所在列</param>
        ///// <param name="projClo">项目名所在的列</param>
        ///// <param name="startRow">开始行</param>
        ///// <param name="endRow">结束行</param>
        ///// <param name="tarPath">目标文件路径</param>
        //public void mergeCells(HSSFWorkbook wb, int sheetIndex, int[] colList, int projClo, int startRow, int endRow, string tarPath)
        //{
        //    string sheetName = wb.GetSheetName(sheetIndex);
        //    mergeCells(wb, sheetName, colList, projClo, startRow, endRow, tarPath);
        //}

        ////合并指定列,按项目名相等
        ////需求是只要项目名相等,则其他需要合并的列数据也一定相等
        ///// <summary>
        ///// 合并指定列,按值相等合并
        ///// </summary>
        ///// <param name="wb">目标workbook</param>
        ///// <param name="sheetName">目标sheet名</param>
        ///// <param name="colList">要合并的单元格所在列</param>
        ///// <param name="projClo">项目名所在的列</param>
        ///// <param name="startRow">开始行</param>
        ///// <param name="endRow">结束行</param>
        ///// <param name="tarPath">目标文件路径</param>
        //public void mergeCells(HSSFWorkbook wb, string sheetName, int[] colList, int projClo, int startRow, int endRow, string tarPath)
        //{
        //    try
        //    {
        //        HSSFSheet sheet = (HSSFSheet)wb.GetSheet(sheetName);
        //        if (sheet == null) return;


        //        #region 检查项目名相等并合并所有目标列
        //        HSSFRow tempHSSFRow = (HSSFRow)sheet.GetRow(startRow);
        //        if(tempHSSFRow==null)
        //            return;
        //        HSSFCell tempHSSFCell = (HSSFCell)tempHSSFRow.GetCell(projClo);
        //        if (tempHSSFCell==null)
        //            return;
        //        string tempCellValue = getCellStringValueAllCase(tempHSSFCell);//最新单元格值
        //        int tempRow = startRow;//最新行号,作为需要合并的起始行
        //        int beforeRow = startRow;//之前的行号,作为需要合并的结束行
        //        for (int j = startRow + 1; j <= endRow; j++)//遍历列的指定行集合
        //        {
        //            HSSFRow tempHSSFRow1 = (HSSFRow)sheet.GetRow(j);
        //            if (tempHSSFRow1 == null)
        //                continue;
        //            HSSFCell tempHSSFCell1 = (HSSFCell)tempHSSFRow1.GetCell(projClo);
        //            if (tempHSSFCell1 == null)
        //                continue;
        //            string nowCellValue = getCellStringValueAllCase(tempHSSFCell1);//目前单元格值

        //            if (tempCellValue.Equals(nowCellValue))//如果相等则之前的行号+1
        //            {
        //                beforeRow++;

        //            }
        //            else//如果不等则合并记录下的单元格区域,并记录新的行号和单元格值
        //            {
        //                if (tempRow < beforeRow)//如果最新行号小于遍历的上一个行号
        //                {
        //                    for (int i = 0; i < colList.Length; i++)//遍历需要合并的列
        //                    {
        //                        //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
        //                        //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
        //                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(tempRow, beforeRow, colList[i], colList[i]));
        //                    }
        //                }
        //                tempCellValue = getCellStringValueAllCase((HSSFCell)sheet.GetRow(j).GetCell(projClo));//更新单元格值
        //                tempRow = j;
        //                beforeRow++;
        //            }


        //        #endregion
        //        }

        //        //保存workbook
        //        saveExcelWithoutAsk(tarPath, wb);
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return;
        //    }
        //}


        ///// <summary>
        ///// 合并指定列,按值相等合并
        ///// </summary>
        ///// <param name="sheetName">目标sheet</param>
        ///// <param name="colList">要合并的单元格所在列</param>
        ///// <param name="startRow">开始行</param>
        ///// <param name="endRow">结束行</param>
        //public HSSFSheet mergeCells(HSSFSheet sheet, int[] colList, int startRow, int endRow)
        //{
        //    try
        //    {
        //        for (int i = 0; i < colList.Length; i++)//遍历需要合并的列
        //        {
        //            #region 检查数值相等并合并当前列
        //            HSSFRow tempHSSFRow = (HSSFRow)sheet.GetRow(startRow);
        //            if (tempHSSFRow == null)
        //                continue;
        //            HSSFCell tempHSSFCell = (HSSFCell)tempHSSFRow.GetCell(colList[i]);
        //            if (tempHSSFCell == null)
        //                continue;
        //            string tempCellValue = getCellStringValueAllCase(tempHSSFCell);//最新单元格值
        //            int tempRow = startRow;//最新行号,作为需要合并的起始行
        //            int beforeRow = startRow;//之前的行号,作为需要合并的结束行
        //            for (int j = startRow + 1; j <= endRow; j++)//遍历列的指定行集合
        //            {
        //                HSSFRow tempHSSFRow1 = (HSSFRow)sheet.GetRow(j);
        //                if (tempHSSFRow1 == null)
        //                    continue;
        //                HSSFCell tempHSSFCell1 = (HSSFCell)tempHSSFRow1.GetCell(colList[i]);
        //                if (tempHSSFCell1 == null)
        //                    continue;
        //                string nowCellValue = getCellStringValueAllCase(tempHSSFCell1);//目前单元格值

        //                if (tempCellValue.Equals(nowCellValue))//如果相等则之前的行号+1
        //                {
        //                    beforeRow++;

        //                }
        //                else//如果不等则合并记录下的单元格区域,并记录新的行号和单元格值
        //                {
        //                    if (tempRow < beforeRow)//如果最新行号小于遍历的上一个行号
        //                    {

        //                        //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
        //                        //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
        //                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(tempRow, beforeRow, colList[i], colList[i]));

        //                    }
        //                    tempCellValue = getCellStringValueAllCase((HSSFCell)sheet.GetRow(j).GetCell(colList[i]));//更新单元格值
        //                    tempRow = j;
        //                    beforeRow++;
        //                }


        //            #endregion
        //            }
        //        }
        //        return sheet;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return sheet;
        //    }
        //}

        ////合并指定行
        ///// <summary>
        ///// 合并指定行,按值相等合并
        ///// </summary>
        ///// <param name="sheet">目标sheet</param>
        ///// <param name="startRow">开始行</param>
        ///// <param name="endRow">结束行</param>
        ///// <param name="startCol">来时列</param>
        ///// <param name="endCol">结束列</param>
        //public HSSFSheet mergeRowCells(HSSFSheet sheet)
        //{
        //    try
        //    {
        //        int startRow = sheet.FirstRowNum;
        //        int endRow = sheet.LastRowNum; ;
        //        for (int i = startRow; i < endRow; i++)//遍历需要合并的行
        //        {
        //            #region 检查数值相等并合并当前列
        //            HSSFRow tempHSSFRow = (HSSFRow)sheet.GetRow(i);
        //            if (tempHSSFRow == null)
        //                continue;
        //            HSSFCell tempHSSFCell = (HSSFCell)tempHSSFRow.GetCell(Int32.Parse(tempHSSFRow.FirstCellNum.ToString()));
        //            if (tempHSSFCell == null)
        //                continue;
        //            string tempCellValue = getCellStringValueAllCase(tempHSSFCell);//最新单元格值
        //            if (tempCellValue == null)
        //            {
        //                tempCellValue="";
        //            }
        //            int tempCol = tempHSSFRow.FirstCellNum;//最新列号,作为需要合并的起始列
        //            int beforeCol = tempHSSFRow.FirstCellNum;//之前的列号,作为需要合并的结束列
        //            for (int j = tempHSSFRow.FirstCellNum + 1; j <= tempHSSFRow.LastCellNum; j++)//遍历列的指定行集合
        //            {
        //                HSSFRow tempHSSFRow1 = (HSSFRow)sheet.GetRow(i);
        //                if (tempHSSFRow1 == null)
        //                    continue;
        //                HSSFCell tempHSSFCell1 = (HSSFCell)tempHSSFRow1.GetCell(j);
        //                if (tempHSSFCell1 == null)
        //                    continue;
        //                string nowCellValue = getCellStringValueAllCase(tempHSSFCell1);//目前单元格值
        //                if (nowCellValue ==null)
        //                {
        //                    nowCellValue = "";
        //                }

        //                if (tempCellValue.Equals(nowCellValue))//如果相等则之前的列号+1
        //                {
        //                    beforeCol++;

        //                }
        //                else//如果不等则合并记录下的单元格区域,并记录新的列号和单元格值
        //                {
        //                    //如果最新列号小于遍历的上一个列号,且单元格值非空
        //                    if (tempCol < beforeCol && !tempCellValue.Equals(""))
        //                    {

        //                        //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
        //                        //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
        //                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, tempCol, beforeCol));

        //                    }
        //                    tempCellValue = getCellStringValueAllCase((HSSFCell)sheet.GetRow(i).GetCell(j));//更新单元格值
        //                    tempCol = j;
        //                    beforeCol++;
        //                }


        //            #endregion
        //            }
        //        }
        //        return sheet;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex, "");
        //        return sheet;
        //    }
        //}
        #endregion

    }
}