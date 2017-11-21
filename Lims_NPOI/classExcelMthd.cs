using System;
using System.IO;
using System.Runtime.InteropServices;
using EXCEL = Microsoft.Office.Interop.Excel;
using WORD = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Collections.Generic;

namespace nsLims_NPOI
{
    class classExcelMthd
    {
        //A4总行高20*41 + 23.25
        //16403, 16865
        //820.25, 首页为14.25pound时843.25, 首页为14.25pound时 833.75
        //手动设置的测试总高,使用不同字体会有不同的总高度,此处使用宋体10号字体
        //private static double PAGE_HEIGHT = 820.15;
        //private static double PAGE_HEIGHT = 842;
        //private static double PAGE_HEIGHT = 820.03;
        //private static double PAGE_HEIGHT = 816;
        private static double PAGE_HEIGHT = 816;
        //最后一页的起始行号
        public int lastPageFirstRow;

        public classExcelMthd()
        {

        }

        //向指定sheet添加图片并保存
        /// <summary>
        /// 向指定sheet添加图片并保存
        /// </summary>
        /// <param name="workbookPath"></param>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="rangeName"></param>
        /// <param name="imagePath"></param>
        /// <param name="imgFlag"></param>
        /// <param name="PicWidth"></param>
        /// <param name="PicHeight"></param>
        /// <returns></returns>
        private string addImageToSheet(string workbookPath, EXCEL.Workbook wb, int sheetIndex, string rangeName,
            string imagePath, string imgFlag, double PicWidth, double PicHeight)
        {
            object missing = System.Reflection.Missing.Value;
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
            if (rangeName.Equals(""))
            {
                return "";
            }
            EXCEL.Range rng = (EXCEL.Range)sheet.get_Range(rangeName, Type.Missing);
            var selectObj = rng.Select();
            float PicLeft, PicTop;    //距离左边距离，顶部距离
            PicTop = Convert.ToSingle(rng.Top);
            PicLeft = Convert.ToSingle(rng.Left);

            var shapes = sheet.Shapes;
            var newShape = shapes.AddPicture(imagePath, Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, (float)PicWidth, (float)PicHeight);

            wb.Save();
            return workbookPath;
            //string newGUID = System.Guid.NewGuid().ToString();
            //string strTargetFile = workbookPath.Substring(0, workbookPath.LastIndexOf("\\") + 1)
            //    + newGUID
            //    + workbookPath.Substring(workbookPath.LastIndexOf("."));
            //wb.SaveAs(strTargetFile, wb.FileFormat, missing, missing, missing, missing, EXCEL.XlSaveAsAccessMode.xlNoChange,
            //    missing, missing, missing, missing, missing);
            //return strTargetFile;
        }


        /// <summary>
        /// 添加图片到指定excel,图片指定大小,使用office的com组件
        /// </summary>
        /// <param name="workbookPath">源工作簿路径</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="imgFlag">图片文件路径</param>
        /// <param name="PicWidth">图片宽度</param>
        /// <param name="PicHeight">图片高度</param>
        /// <returns></returns>
        public bool addImage2Excel_byOffice(string workbookPath, int sheetIndex, string imagePath, string imgFlag, double PicWidth, double PicHeight)
        {
            bool flag = true;
            if (!File.Exists(workbookPath))
            {
                return false;
            }
            //object missing = Type.Missing;
            object missing = System.Reflection.Missing.Value;
            string strTargetFile = "";
            //按照标记找到range单元格名
            string rangeName = new classLims_NPOI().getExcelRangeByFlag(workbookPath, sheetIndex, imgFlag);
            EXCEL.ApplicationClass excel = null;
            EXCEL.Workbook wb = null;
            EXCEL.Workbooks workBooks = null;
            try
            {
                excel = new EXCEL.ApplicationClass();
                excel.DisplayAlerts = false;
                workBooks = excel.Workbooks;
                wb = workBooks.Open(workbookPath, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing);
                //实例化Sheet后,释放Excel进程就会失败
                //对于sheet的操作必须放在新的方法中,接口层级为Workbook
                strTargetFile = addImageToSheet(workbookPath, wb, sheetIndex + 1, rangeName, imagePath, imgFlag, PicWidth, PicHeight);

            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                flag = false;
            }
            finally
            {
                if (wb != null)
                {
                    //wb.Close(false, missing, false);
                    wb.Close(false, missing, missing);
                    int i = Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (workBooks != null)
                {
                    workBooks.Close();
                    int i = Marshal.ReleaseComObject(workBooks);
                    workBooks = null;
                }
                if (excel != null)
                {
                    excel.Quit();
                    int i = Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
            //string strOldFileName = workbookPath.Substring(workbookPath.LastIndexOf("\\") + 1,
            //    workbookPath.Length - workbookPath.LastIndexOf("\\") - 1);
            //if (File.Exists(strTargetFile))
            //{
            //    File.Delete(workbookPath);
            //    Computer MyComputer = new Computer();
            //    MyComputer.FileSystem.RenameFile(strTargetFile, strOldFileName);
            //    flag = true;
            //}
            //else
            //{
            //    flag = false;
            //}
            return flag;

        }


        /// <summary>
        /// 添加图片到指定excel,图片指定大小,使用office的com组件, 批量添加;
        ///      在签名图片高度保持0.9cm( 25.4磅)的情况下,要求图片宽高比不能高于黄金比例(0.618),否则将盖不住标记字符串
        /// </summary>
        /// <param name="workbookPath">源工作簿路径</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="dArray">图片添加位置标记和图片文件路径数组</param>
        /// <param name="PicWidth">图片宽度,无实际意义,实际值是按高度和图片比例计算出来的</param>
        /// <param name="PicHeight">图片高度</param>
        /// <returns></returns>
        public bool addImagesToExcel_byOffice(string workbookPath, int sheetIndex, object[] dArray, double PicWidth, double PicHeight)
        {
            bool flag = true;
            try
            {
                if (!File.Exists(workbookPath))
                {
                    return false;
                }
                //object missing = Type.Missing;
                object missing = System.Reflection.Missing.Value;
                List<string> rangeNameList = new List<string>();
                //按照标记找到位置
                Dictionary<string, string> dictionary = classLims_NPOI.dArray2Dictionary(dArray);
                foreach (var oneMapPoint in dictionary)
                {
                    string key = oneMapPoint.Key.ToString();
                    string value = oneMapPoint.Value.ToString();
                    //按照标记找到range单元格名
                    string rangeName = new classLims_NPOI().getExcelRangeByFlag(workbookPath, sheetIndex, key);
                    if (rangeName == null || rangeName == "")
                    {
                        rangeNameList.Add("");
                    }
                    else
                    {
                        rangeNameList.Add(rangeName);
                    }
                }


                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook wb = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    wb = workBooks.Open(workbookPath, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    //实例化Sheet后,释放Excel进程就会失败
                    //对于sheet的操作必须放在新的方法中,接口层级为Workbook
                    //按照标记找到位置并插入图片
                    int i = 0;
                    foreach (var oneMapPoint in dictionary)
                    {
                        string key = oneMapPoint.Key.ToString();
                        string value = oneMapPoint.Value.ToString();

                        if (rangeNameList[i] == null || rangeNameList[i] == "")
                        {
                            classLims_NPOI.WriteLog("标记字符串:" + key + " 未检测到!", "");
                            i++;
                            if (i > rangeNameList.Count)
                                break;
                            continue;
                        }
                        if (value == null || value == "" || !File.Exists(value))
                        {
                            classLims_NPOI.WriteLog("签名文件:" + value + " 不存在!", "");
                            i++;
                            if (i > rangeNameList.Count)
                                break;
                            continue;
                        }

                        //应该是签名图片高度固定,但宽度会等比例缩放
                        System.Drawing.Image image = System.Drawing.Image.FromFile(value);
                        var imageWidth = image.Width;
                        var imageHeight = image.Height;
                        PicWidth = PicHeight * imageWidth / imageHeight;
                        addImageToSheet(workbookPath, wb, sheetIndex + 1, rangeNameList[i], value, key, PicWidth, PicHeight);
                        i++;
                        if (i > rangeNameList.Count)
                            break;
                    }
                    wb.Save();
                    //strTargetFile = addImageToSheet(workbookPath, wb, sheetIndex + 1, rangeName, imagePath, imgFlag, PicWidth, PicHeight);

                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (wb != null)
                    {
                        //wb.Close(false, missing, false);
                        wb.Close(false, missing, missing);
                        int i = Marshal.ReleaseComObject(wb);
                        wb = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        int i = Marshal.ReleaseComObject(workBooks);
                        workBooks = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        int i = Marshal.ReleaseComObject(excel);
                        excel = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
            }
            catch (Exception e)
            {
                classLims_NPOI.WriteLog(e, "");
                flag = false;
            }
            return flag;

        }

        /// <summary>
        /// 添加图片到指定excel,图片使用原始尺寸,使用office的com组件, 批量添加,
        ///         图片宽高磅数设为-1时会使用原始尺寸
        /// </summary>
        /// <param name="workbookPath">源工作簿路径</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="dArray">图片添加位置标记和图片文件路径数组</param>
        /// <returns></returns>
        public bool addImagesToExcel_byOffice(string workbookPath, int sheetIndex, object[] dArray)
        {
            bool flag = true;
            try
            {
                if (!File.Exists(workbookPath))
                {
                    return false;
                }
                //object missing = Type.Missing;
                object missing = System.Reflection.Missing.Value;
                List<string> rangeNameList = new List<string>();
                //按照标记找到位置
                Dictionary<string, string> dictionary = classLims_NPOI.dArray2Dictionary(dArray);
                foreach (var oneMapPoint in dictionary)
                {
                    string key = oneMapPoint.Key.ToString();
                    string value = oneMapPoint.Value.ToString();
                    //按照标记找到range单元格名
                    string rangeName = new classLims_NPOI().getExcelRangeByFlag(workbookPath, sheetIndex, key);
                    if (rangeName == null || rangeName == "")
                    {
                        rangeNameList.Add("");
                    }
                    else
                    {
                        rangeNameList.Add(rangeName);
                    }
                }


                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook wb = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    wb = workBooks.Open(workbookPath, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    //实例化Sheet后,释放Excel进程就会失败
                    //对于sheet的操作必须放在新的方法中,接口层级为Workbook
                    //按照标记找到位置并插入图片
                    int i = 0;
                    foreach (var oneMapPoint in dictionary)
                    {
                        string key = oneMapPoint.Key.ToString();
                        string value = oneMapPoint.Value.ToString();

                        if (rangeNameList[i] == null || rangeNameList[i] == "")
                        {
                            classLims_NPOI.WriteLog("标记字符串:" + key + " 未检测到!", "");
                            i++;
                            if (i > rangeNameList.Count)
                                break;
                            continue;
                        }
                        if (value == null || value == "" || !File.Exists(value))
                        {
                            classLims_NPOI.WriteLog("签名文件:" + value + " 不存在!", "");
                            i++;
                            if (i > rangeNameList.Count)
                                break;
                            continue;
                        }

                        addImageToSheet(workbookPath, wb, sheetIndex + 1, rangeNameList[i], value, key, -1, -1);
                        i++;
                        if (i > rangeNameList.Count)
                            break;
                    }
                    wb.Save();
                    //strTargetFile = addImageToSheet(workbookPath, wb, sheetIndex + 1, rangeName, imagePath, imgFlag, PicWidth, PicHeight);

                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (wb != null)
                    {
                        //wb.Close(false, missing, false);
                        wb.Close(false, missing, missing);
                        int i = Marshal.ReleaseComObject(wb);
                        wb = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        int i = Marshal.ReleaseComObject(workBooks);
                        workBooks = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        int i = Marshal.ReleaseComObject(excel);
                        excel = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
            }
            catch (Exception e)
            {
                classLims_NPOI.WriteLog(e, "");
                flag = false;
            }
            return flag;

        }


        public static bool createNew03Excel(string strSourceFile)
        {
            bool flag = true;
            if (true)
            {
                if (File.Exists(strSourceFile))
                {
                    File.Delete(strSourceFile);
                }
                var Ext = strSourceFile.Substring(strSourceFile.LastIndexOf("."));
                if (Ext.ToUpper() != ".XLS")
                {
                    //throw new Exception("文件格式不符合要求!");
                    classLims_NPOI.WriteLog("文件格式不符合要求!", "");
                    flag = false;
                    return flag;
                }
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Add(missing);
                    workBook.SaveAs(strSourceFile, EXCEL.XlFileFormat.xlExcel8, null, null, false, false, EXCEL.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return flag;
        }

        /// <summary>
        /// 处理2页之间的合并单元格,上下单独合并
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="endCol"></param>
        /// <param name="colseq">需要合并的列</param>
        public void dealMergedAreaInPages_new(EXCEL.Workbook wb, int sheetIndex, int endCol, int[] colseq)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                int endRow = sheet.UsedRange.Rows.Count;

                System.Drawing.Point p = selectPosition(sheet, "检测项目", endRow, endCol);
                int testNoIndex = p.Y;
                //列号取检测项目标题所在行的单元格数
                //int maxColIndex = sheet.Range["IV" + testNoIndex].End[EXCEL.XlDirection.xlToLeft].Column;
                int maxColIndex = endCol;

                //获取每页第一行的索引
                List<int> firstPageRowIndex = getNewPageFirstRow(sheet);
                //记录下最后一页的起始行号,如果页数不超过一页,则应该返回表头下第一行行号
                if (firstPageRowIndex.Count == 0)
                {
                    int[] iRange = getRepeatingRowsRange(sheet, endCol);
                    if (iRange == null)
                    {
                        this.lastPageFirstRow = 1;
                    }
                    else
                    {
                        this.lastPageFirstRow = iRange[1] + 1;
                    }
                }
                else
                {
                    this.lastPageFirstRow = firstPageRowIndex[firstPageRowIndex.Count - 1];
                }

                foreach (int i in firstPageRowIndex)
                {

                    if (i <= endRow && i > p.X)
                    {
                        #region 不能按照检测项目相等判断是否合并, 因为不同检测项目前可能是相同的子样名称
                        ////如果检测项目相等,则拆分
                        //string boforeCellValue = getMergerCellValue(sheet, i - 1, testNoIndex);
                        //string cellValue = getMergerCellValue(sheet, i, testNoIndex);
                        //if (boforeCellValue.Equals("") || cellValue.Equals(""))
                        //{
                        //    ;
                        //}
                        //else if (boforeCellValue.Equals(cellValue))
                        //{
                        //    dealMergedBetweenPages(sheet, i, maxColIndex, colseq);
                        //}
                        #endregion
                        dealMergedBetweenPages(sheet, i, maxColIndex, colseq);
                    }
                    //插入分页符
                    //((EXCEL.Range)sheet.Rows[20]).PageBreak = 1;

                }
                classLims_NPOI.WriteLog("dealMergedAreaInPages_new success", "");
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }

        /// <summary>
        /// 处理2页之间的合并单元格,上下单独合并,不需要合并列的列不用处理
        /// </summary>
        /// <param name="sheet">工作表对象</param>
        /// <param name="rowIndex">页首行行号</param>
        /// <param name="endCol">最大行</param>
        /// <param name="colseq">需要合并的列</param>
        private void dealMergedBetweenPages(EXCEL.Worksheet sheet, int rowIndex, int endCol, int[] colseq)
        {
            try
            {
                int i = 0;
                while (i < colseq.Length && i <= endCol)
                {
                    //classLims_NPOI.WriteLog("当前I值:" + i.ToString(), "");
                    EXCEL.Range cell = (EXCEL.Range)sheet.Cells[rowIndex, colseq[i]];
                    if (cell == null) continue;
                    if ((bool)cell.MergeCells)
                    {

                        //先获取合并区域,不拆分,需要处理跨页再拆分
                        int[] mergedArea = getMergedArea(sheet, rowIndex, colseq[i], false);

                        //合并区域的起始行大于页起始行可不用处理
                        if (mergedArea[0] >= rowIndex)
                        {
                            //不跨页处理的不用拆分
                            i++;
                            continue;
                        }
                        else
                        {
                            //获取合并区域的值,拆分合并后需要重新赋值
                            var cellValue = getMergerCellValue(sheet, rowIndex, colseq[i]);

                            //此处代表需要跨页拆分
                            var mgIndexT = sheet.get_Range(sheet.Cells[mergedArea[0], mergedArea[1]], sheet.Cells[mergedArea[2], mergedArea[3]]);
                            mgIndexT.UnMerge();
                            mgIndexT.Value = cellValue;//拆分后必须重设所有合并区域值,否则多出的单元格将为空


                            //合并上一页
                            var mgIndex1 = sheet.get_Range(sheet.Cells[mergedArea[0], mergedArea[1]], sheet.Cells[rowIndex - 1, mergedArea[3]]);
                            mgIndex1.Merge(Missing.Value);
                            //设置边框为全框线
                            mgIndex1.Borders.LineStyle = 1;

                            //合并下一页
                            var mgIndex2 = sheet.get_Range(sheet.Cells[rowIndex, mergedArea[1]], sheet.Cells[mergedArea[2], mergedArea[3]]);
                            mgIndex2.Merge(Missing.Value);
                            //设置边框为全框线
                            mgIndex2.Borders.LineStyle = 1;

                            #region 应该跳到合并区域之外,
                            //检查合并区域的结束列,找到比它大的第一个合并列数组索引
                            bool jump = false;//是否跳过
                            for (int j = i + 1; j < colseq.Length; j++)
                            {
                                //mergedArea[3]
                                if (colseq[j] > mergedArea[3])
                                {
                                    i = j;
                                    jump = true;
                                    break;
                                }
                            }
                            if (jump == true)
                            {
                                continue;
                            }
                            else
                            {
                                i++;
                                continue;
                            }
                            #endregion
                        }
                    }
                    else
                    {
                        i++;
                        continue;
                    }
                }
            }
            catch (Exception e)
            {
                classLims_NPOI.WriteLog(e, "");
            }


        }

        /// <summary>
        /// excel刷新
        /// </summary>
        /// <param name="strSourceFile">excel文件绝对路径</param>
        /// <returns></returns>
        public static bool excelRefresh(string strSourceFile)
        {
            bool flag = true;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    //EXCEL.XlFileFormat.xlAddIn: xls
                    //XlFileFormat.xlOpenXMLWorkbook:          xlsx                    
                    //workBook.SaveAs(targetFile, workBook.FileFormat, missing, missing, missing, missing, EXCEL.XlSaveAsAccessMode.xlNoChange,
                    //    missing, missing, missing, missing, missing);

                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return flag;
        }

        /// <summary>
        /// excel刷新
        /// </summary>
        /// <param name="strSourceFile">excel文件绝对路径</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="startRow">起始行</param>
        /// <param name="endRow">结束行</param>
        /// <param name="updHeight">增加的高度</param>
        /// <returns></returns>
        public static bool excelRefreshAndUpdateRowHeight(string strSourceFile, int sheetIndex, int startRow, int endRow, double updHeight)
        {
            bool flag = true;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    //EXCEL.XlFileFormat.xlAddIn: xls
                    //XlFileFormat.xlOpenXMLWorkbook:          xlsx                    
                    //workBook.SaveAs(targetFile, workBook.FileFormat, missing, missing, missing, missing, EXCEL.XlSaveAsAccessMode.xlNoChange,
                    //    missing, missing, missing, missing, missing);

                    updateSheesRowHeight(workBook, sheetIndex, startRow, endRow, updHeight);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return flag;
        }


        //在另外一个sheet里面利用单元格换行和自适应高度的特性,将一个试验单元格宽度设置成实际跨列单元格的宽度,
        //然后将需要输入的字符放入该试验单元格,取得高度返回给实际跨列单元格就可以了.
        private static double getCellAutoHeight(EXCEL.Workbook wb, int sheetIndex, int row, int col)
        {
            double iHeight = -1;
            object missing = Type.Missing;
            try
            {
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

                #region 计算合并区域总列宽
                int[] mergedArea = getMergedArea(sheet, row, col, false);
                //如果是合并单元格靠后的单元格,代表已经计算过,可以直接忽略
                if (mergedArea[1] < col) return iHeight;
                EXCEL.Range mergedCell = sheet.get_Range(sheet.Cells[mergedArea[0], mergedArea[1]], sheet.Cells[mergedArea[2], mergedArea[3]]);
                mergedCell.Application.DisplayAlerts = false;
                double colWidth = 0;
                for (int i = mergedArea[1]; i <= mergedArea[3]; i++)
                {
                    colWidth += (double)((EXCEL.Range)sheet.Columns[i]).ColumnWidth;

                }
                #endregion
                EXCEL.Range cell = (EXCEL.Range)sheet.Cells[row, col];


                wb.Sheets.Add(missing, sheet, 1, missing);
                EXCEL.Worksheet wkSheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex + 1];
                EXCEL.Range wkCell = (EXCEL.Range)wkSheet.Cells[1, 1];
                wkCell.ColumnWidth = colWidth;
                wkCell.Style = cell.Style;//先同步单元格格式
                wkCell.Font.FontStyle = cell.Font.FontStyle;
                wkCell.Font.Size = cell.Font.Size;
                wkCell.Font.Name = cell.Font.Name;
                wkCell.WrapText = true;//再设为自动换行
                wkCell.Value = getMergerCellValue(sheet, row, col);
                //获取自适应后的行高
                iHeight = (double)wkCell.RowHeight;
                ((EXCEL.Worksheet)wb.Sheets[sheetIndex + 1]).Delete();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
            return iHeight;
        }

        //获取单元格所需行高
        private static double getCellRowHeight(EXCEL.Workbook wb, int sheetIndex, int row, int col)
        {
            double height = 0;
            try
            {
                object missing = System.Reflection.Missing.Value;
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

                //获取需要合并的单元格的范围
                //EXCEL.Range cell = sheet.get_Range(sheet.Cells[startRow, startCol], sheet.Cells[startRow, startCol]);
                int[] mergedArea = getMergedArea(sheet, row, col, false);
                EXCEL.Range mergedCell = sheet.get_Range(sheet.Cells[mergedArea[0], mergedArea[1]], sheet.Cells[mergedArea[2], mergedArea[3]]);
                mergedCell.Application.DisplayAlerts = false;
                //
                string cellValue = getMergerCellValue(sheet, row, col);
                int valueLength = getMaxStringCharLength(cellValue);
                double colWidth = 0;
                for (int i = mergedArea[1]; i <= mergedArea[3]; i++)
                {
                    colWidth += (double)((EXCEL.Range)sheet.Columns[i]).ColumnWidth;

                }
                double fontSize = (double)mergedCell.Font.Size;//10号字体通项公式为0.875+(n)*0.875=(n+1)*0.875
                double autoColWidth;
                int cellRow;
                if (fontSize == 10)
                {
                    //autoColWidth = (valueLength + 1) * 0.875;
                    autoColWidth = (valueLength + 0) * 0.875;
                    //列宽精度为0.001, 需要向上取整
                    cellRow = (int)(autoColWidth / colWidth + 0.999);
                    if (cellRow == 1)
                    {
                        height = 14.25;
                    }
                    else
                    {
                        height = 12 * cellRow;
                    }
                }
                else
                {
                    classLims_NPOI.WriteLog("字体大小不是10号.", "");
                    height = -1;
                }
                //mergedCell.Application.DisplayAlerts = true;

                return height;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return -1;
            }
        }

        //获取单元格所需行高
        private static double getCellRowHeight(string strSourceFile, int sheetIndex, int row, int col)
        {
            double height = 0;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    height = getCellRowHeight(workBook, sheetIndex, row, col);
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return height;
        }

        /// <summary>
        /// 获取字符串长度,英文占1个,中文占2个,判断方式为ASCII
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static int getMaxStringCharLength(string str)
        {
            int length = 0;
            for (int i = 0; i < str.Length; i++)
            {
                if ((int)str[i] <= 127)
                    length++;
                else
                    length = length + 2;
            }
            return length;
        }

        //获取行高,按最大值
        /// <summary>
        /// 获取行高,按最大值
        /// </summary>
        /// <param name="wb">目标工作簿</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="row">行号</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <returns>行高</returns>
        private static double getRowHeight(EXCEL.Workbook wb, int sheetIndex, int row, int startCol, int endCol)
        {
            double maxRowHeight = 0;
            try
            {
                object missing = System.Reflection.Missing.Value;
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                //EXCEL.Range rowRange = (EXCEL.Range)sheet.Rows[row];

                for (int i = startCol; i <= endCol; i++)
                {
                    double tempRowHeight = getCellAutoHeight(wb, sheetIndex, row, i);
                    if (maxRowHeight < tempRowHeight)
                        maxRowHeight = tempRowHeight;
                }
                return maxRowHeight;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return -1;
            }
        }

        //获取行高,按最大值
        /// <summary>
        /// 获取行高,按最大值
        /// </summary>
        /// <param name="strSourceFile">目标文档</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="row">行号</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <returns>行高</returns>
        public static double getRowHeight(string strSourceFile, int sheetIndex, int row, int startCol, int endCol)
        {
            double height = 0;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    height = getRowHeight(workBook, sheetIndex, row, startCol, endCol);
                    //workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return height;
        }

        //采用office组件合并单元格,避免合并后无法适应行高
        public static void mergeRowCells_ByOffice(EXCEL.Workbook wb, int sheetIndex, int startRow, int endRow, int startCol, int endCol)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

                //获取需要合并的单元格的范围
                EXCEL.Range rangeProgram = sheet.get_Range(sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]);
                rangeProgram.Application.DisplayAlerts = false;
                rangeProgram.Merge(Missing.Value);
                rangeProgram.Application.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }

        /// <summary>
        /// 采用office组件合并单元格,避免合并后无法适应行高
        /// </summary>
        /// <param name="strSourceFile"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startRow">起始行</param>
        /// <param name="endRow">结束行</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <returns></returns>
        public static bool mergeRowCells_ByOffice(string strSourceFile, int sheetIndex, int startRow, int endRow, int startCol, int endCol)
        {
            bool flag = true;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    mergeRowCells_ByOffice(workBook, sheetIndex, startRow, endRow, startCol, endCol);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return flag;
        }

        //获得单元格的合并区域
        /// <summary>
        /// 获得单元格的合并区域, 起始行, 起始列, 结束行, 结束列
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="nRow">行索引</param>
        /// <param name="nCol">列索引</param>
        /// <param name="isUnMegre">是否取消合并</param>
        /// <returns></returns>
        public static int[] getMergedArea(EXCEL.Worksheet sheet, int nRow, int nCol, bool isUnMegre)
        {
            int[] pa = { 0, 0, 0, 0 };
            try
            {

                //获取需要合并的单元格的范围
                EXCEL.Range rangeCell = (EXCEL.Range)sheet.Cells[nRow, nCol];
                rangeCell.Application.DisplayAlerts = false;
                var ma = rangeCell.MergeArea;


                pa[0] = ma.Row;
                pa[1] = ma.Column;
                pa[2] = pa[0] + ma.Rows.Count - 1;
                pa[3] = pa[1] + ma.Columns.Count - 1;

                //取消合并区域
                if (isUnMegre)
                {
                    //EXCEL.Range leftTopCell = (EXCEL.Range)sheet.Cells[ma.Row, ma.Column];
                    string leftTopCellValue = getMergerCellValue(sheet, ma.Row, ma.Column);
                    ma.UnMerge();
                    ma.Value = leftTopCellValue;
                }

                return pa;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return pa;
            }
        }

        //获得单元格的合并区域
        /// <summary>
        /// 获得单元格的合并区域, 起始行, 起始列, 结束行, 结束列
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="nRow"></param>
        /// <param name="nCol"></param>
        /// <param name="isUnMegre">是否取消合并</param>
        /// <returns></returns>
        public static int[] getMergedArea(EXCEL.Workbook wb, int sheetIndex, int nRow, int nCol, bool isUnMegre)
        {
            int[] pa = { 0, 0, 0, 0 };
            try
            {
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                pa = getMergedArea(sheet, nRow, nCol, isUnMegre);

                return pa;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return pa;
            }
        }


        //获得单元格的合并区域
        /// <summary>
        /// 获得单元格的合并区域:____起始行, 起始列, 结束行, 结束列
        /// </summary>
        /// <param name="strSourceFile"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="nRow"></param>
        /// <param name="nCol"></param>
        /// <param name="isUmMerge">是否取消合并</param>
        /// <returns></returns>
        public static int[] getMergedArea(string strSourceFile, int sheetIndex, int nRow, int nCol, bool isUmMerge)
        {
            int[] pa = { 0, 0, 0, 0 };
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    pa = getMergedArea(workBook, sheetIndex, nRow, nCol, isUmMerge);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return pa;
        }

        //获取合并后单元格的值,为左上角单元格的值
        /// <summary>
        /// 获取合并后单元格的值,为左上角单元格的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        /// <returns></returns>
        private static string getMergerCellValue(EXCEL.Worksheet sheet, int row, int col)
        {
            try
            {
                //索引超标则记录行列索引
                int endRow = sheet.UsedRange.Rows.Count;
                int endCol = sheet.UsedRange.Columns.Count;
                if (row > endRow || col > endCol)
                {
                    string errorMessage = "";
                    errorMessage += "获取单元格值索引超出范围:[最大行,最大列]=["
                        + endRow.ToString() + "," + endCol.ToString() + "]; [实际行,实际列]=["
                        + row.ToString() + "," + col.ToString() + "]";
                    classLims_NPOI.WriteLog(errorMessage, "");
                    return "";
                }
                EXCEL.Range cell = (EXCEL.Range)sheet.Cells[row, col];
                //EXCEL.Range cell1 = (EXCEL.Range)sheet.Cells[4, 6];
                if ((bool)cell.MergeCells)
                {
                    var ma = cell.MergeArea;
                    EXCEL.Range leftTopCell = sheet.get_Range(sheet.Cells[ma.Row, ma.Column], sheet.Cells[ma.Row, ma.Column]);
                    if (leftTopCell == null)
                        return "";
                    var cv = leftTopCell.Value;
                    if (cv == null)
                        return "";
                    return cv.ToString();
                }
                else
                {
                    var cv = cell.Value;
                    if (cv == null)
                        return "";
                    return cv.ToString();
                }
            }
            catch (Exception e)
            {
                classLims_NPOI.WriteLog(e, "");
                return "";
            }
        }

        #region 废弃的,使用计算 获取默认分页位置
        ////获取sheet每一页的第一行索引,第一页除外
        ///// <summary>
        ///// 获取sheet每一页的第一行索引,第一页除外
        ///// </summary>
        ///// <param name="sheet"></param>
        ///// <returns></returns>
        //private List<int> getNewPageFirstRow(EXCEL.Worksheet sheet, int endRow, int endCol)
        //{
        //    List<int> arrayFr = new List<int>();
        //    try
        //    {
        //        int startHeadRow, endHeadRow;
        //        int[] iRange = getRepeatingRowsRange(sheet, endCol);
        //        if (iRange == null)
        //        {
        //            startHeadRow = 0;
        //            endHeadRow = 0;
        //        }
        //        else
        //        {
        //            startHeadRow = iRange[0];
        //            endHeadRow = iRange[1];
        //        }
        //        int fRow = endHeadRow + 1;
        //        //int lRow = sheet.Range["A65535"].End[EXCEL.XlDirection.xlUp].Row;
        //        int lRow = endRow;

        //        double totalH = 0; //总行高
        //                           //高度换算 1厘米＝28.346456692913389磅
        //        double headH = 0; //表头高度
        //        headH = sheet.PageSetup.TopMargin + sheet.PageSetup.BottomMargin;
        //        if (iRange == null)
        //        {
        //            headH += 0;
        //        }
        //        else
        //        {
        //            ////不使用区域提取总行高,会不准确;改为使用遍历行高的方式
        //            //double tempTitleHeight = 0;
        //            //for(int i= iRange[0]; i<= iRange[1]; i++)
        //            //{
        //            //    EXCEL.Range row = (EXCEL.Range)sheet.Rows[i];
        //            //    tempTitleHeight += (double)row.Height;
        //            //}
        //            //headH += tempTitleHeight;
        //            headH += (double)sheet.get_Range(sheet.Cells[iRange[0], iRange[2]], sheet.Cells[iRange[1], iRange[3]]).Height;
        //        }

        //        for (int i = fRow; i <= lRow; i++)
        //        {
        //            //classLims_NPOI.WriteLog("i="+i,"");
        //            EXCEL.Range row = (EXCEL.Range)sheet.Rows[i];
        //            if (row == null) continue;

        //            double tempH;
        //            if ((bool)((EXCEL.Range)sheet.Rows[i]).Hidden == true)
        //            {
        //                tempH = 0;
        //            }
        //            else
        //            {
        //                tempH = (double)((EXCEL.Range)sheet.Rows[i]).Height;
        //            }

        //            if (System.Convert.ToInt32(totalH + tempH) >= System.Convert.ToInt32(PAGE_HEIGHT - (1 * headH)))//超过一页
        //            {
        //                arrayFr.Add(i);
        //                totalH = tempH;
        //            }
        //            else
        //            {
        //                totalH = totalH + tempH;
        //            }
        //        }
        //        return arrayFr;
        //    }
        //    catch (Exception ex)
        //    {
        //        classLims_NPOI.WriteLog(ex, "");
        //        return arrayFr;
        //    }

        //}


        ////获取sheet每一页的第一行索引,第一页除外
        ///// <summary>
        ///// 获取sheet每一页的第一行索引,第一页除外
        ///// </summary>
        ///// <param name="sheet"></param>
        ///// <returns></returns>
        //private List<int> getNewPageFirstRow(EXCEL.Worksheet sheet, int endRow, int endCol)
        //{
        //    List<int> arrayFr = new List<int>();
        //    try
        //    {
        //        //获取默认行分页符集合
        //        var hpb = sheet.HPageBreaks;
        //        for (int i = 0; i < hpb.Count; i++)
        //        {
        //            var hpbRange = hpb[i].Location;
        //            arrayFr.Add(hpbRange.Row);
        //        }
        //        return arrayFr;
        //    }
        //    catch (Exception ex)
        //    {
        //        classLims_NPOI.WriteLog(ex, "");
        //        return arrayFr;
        //    }

        //}
        #endregion

        //获取sheet每一页的第一行索引,第一页除外
        /// <summary>
        /// 获取sheet每一页的第一行索引,第一页除外
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static List<int> getNewPageFirstRow(EXCEL.Worksheet sheet)
        {
            List<int> arrayFr = new List<int>();
            try
            {
                //获取默认行分页符集合
                EXCEL.HPageBreaks hpb = sheet.HPageBreaks;

                for (int i = 1; i <= hpb.Count; i++)
                {
                    var hPageBreak = hpb[i];
                    if (hPageBreak != null)
                    {
                        var hpbRange = hPageBreak.Location;
                        arrayFr.Add(hpbRange.Row);
                    }
                }
                return arrayFr;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return arrayFr;
            }

        }

        //获取sheet每一页的第一行索引,第一页除外
        /// <summary>
        /// 获取sheet每一页的第一行索引,第一页除外
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<int> getNewPageFirstRow(string sourceFile, int sheetIndex)
        {
            List<int> arrayFr = new List<int>();
            object missing = System.Reflection.Missing.Value;
            EXCEL.ApplicationClass excel = null;
            EXCEL.Workbook wb = null;
            EXCEL.Workbooks workBooks = null;
            try
            {
                excel = new EXCEL.ApplicationClass();
                excel.DisplayAlerts = false;
                workBooks = excel.Workbooks;
                wb = workBooks.Open(sourceFile, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing);
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                arrayFr = getNewPageFirstRow(sheet);
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
            finally
            {
                if (wb != null)
                {
                    //wb.Close(false, missing, false);
                    wb.Close(false, missing, missing);
                    int i = Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (workBooks != null)
                {
                    workBooks.Close();
                    int i = Marshal.ReleaseComObject(workBooks);
                    workBooks = null;
                }
                if (excel != null)
                {
                    excel.Quit();
                    int i = Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
            return arrayFr;
        }

        /// <summary>
        /// 返回sheet重复区域,分别为起始行号,结束行号,起始列号,结束列号
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>整型数组,分别为起始行号,结束行号,起始列号,结束列号</returns>
        private static int[] getRepeatingRowsRange(EXCEL.Worksheet sheet, int endCol)
        {
            int printCount = sheet.PrintedCommentPages;
            string pntRow = sheet.PageSetup.PrintTitleRows;//"$1:$2"
            //classLims_NPOI.WriteLog("顶端标题行:"+ pntRow, "");
            string pntCol = sheet.PageSetup.PrintTitleColumns;
            if (pntRow == null || pntRow.Equals("")) { return null; }
            int[] ir = new int[] { -1, -1, -1, -1 };
            ir[0] = Int32.Parse(pntRow.Substring(1, pntRow.IndexOf(":") - 1));
            ir[1] = Int32.Parse(pntRow.Substring(pntRow.IndexOf(":") + 2));

            if (pntCol == null || pntCol.Equals(""))
            {
                ir[2] = 1;
                ir[3] = endCol;
            }
            else
            {
                ir[2] = Int32.Parse(pntCol.Substring(1, pntCol.IndexOf(":") - 1));
                ir[3] = Int32.Parse(pntCol.Substring(pntCol.IndexOf(":") + 2));
            }
            return ir;

        }

        /// <summary>
        /// 返回sheet重复区域,分别为起始行号,结束行号,起始列号,结束列号
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sheetIndex">表索引</param>
        /// <param name="endCol">最大列号</param>
        /// <returns>整型数组,分别为起始行号,结束行号,起始列号,结束列号</returns>
        private int[] getRepeatingRowsRange(EXCEL.Workbook wb, int sheetIndex, int endCol)
        {
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

            return getRepeatingRowsRange(sheet, endCol);

        }

        /// <summary>
        /// 返回sheet重复区域,分别为起始行号,结束行号,起始列号,结束列号
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>整型数组,分别为起始行号,结束行号,起始列号,结束列号</returns>
        public int[] getRepeatingRowsRange(string strSourceFile, int sheetIndex, int endCol)
        {
            int[] ir = new int[] { -1, -1, -1, -1 };
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    ir = getRepeatingRowsRange(workBook, sheetIndex, endCol);
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return ir;

        }

        /// <summary>
        /// 获取sheet页数,通过垂直分页符个数判断
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public static int getSheetPageCount(EXCEL.Workbook wb, int sheetIndex)
        {
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

            var hpb = sheet.HPageBreaks;
            int hpbCount = sheet.HPageBreaks.Count;
            return hpbCount + 1;

        }

        /// <summary>
        /// 合并指定列,按值相等合并
        /// </summary>
        /// <param name="sheetName">目标sheet</param>
        /// <param name="colList">要合并的单元格所在列</param>
        /// <param name="startRow">开始行</param>
        /// <param name="endCol">结束列</param>
        public void mergeCells(EXCEL.Workbook wb, int sheetIndex, int[] colList, int startRow, int endCol)
        {
            try
            {
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                int endRow = sheet.UsedRange.Rows.Count;
                int TESTNO_colIndex = selectPosition(sheet, "检测项目", endRow, endCol).Y;//检测项目所在列号
                for (int i = 0; i < colList.Length; i++)//遍历需要合并的列
                {
                    #region 检查数值相等并合并当前列

                    EXCEL.Range tempRowRange = (EXCEL.Range)sheet.Rows[startRow];
                    EXCEL.Range tempCell = (EXCEL.Range)tempRowRange.Cells[colList[i]];
                    if (tempRowRange == null || tempCell == null) continue;
                    string tempCellValue = getMergerCellValue(sheet, startRow, colList[i]);

                    //最新检测项值
                    string tempTESTNO = getMergerCellValue(sheet, startRow, TESTNO_colIndex);

                    int tempRow = startRow;//最新行号,作为需要合并的起始行
                    int beforeRow = startRow;//之前的行号,作为需要合并的结束行
                    for (int j = startRow + 1; j <= endRow; j++)//遍历列的指定行集合
                    {
                        //目前单元格值
                        string nowCellValue = getMergerCellValue(sheet, j, colList[i]);

                        //目前检测项值
                        string nowTESTNO = getMergerCellValue(sheet, j, TESTNO_colIndex);

                        //如果相等则之前的行号+1,需要根据检测项相等判定合并
                        //除序号列外,在检测项目之前的列不用考虑检测项目是否相等再合并
                        if (colList[i] > 1 && colList[i] < TESTNO_colIndex && tempCellValue.Equals(nowCellValue))
                        {
                            //classLims_NPOI.WriteLog("当前索引" + colList[i].ToString(), "");
                            //classLims_NPOI.WriteLog("当前值"+ nowCellValue, "");
                            beforeRow++;

                        }
                        //序号和检测项目之后的列,要按照先判断项目是否相等
                        else if (tempCellValue.Equals(nowCellValue) && tempTESTNO.Equals(nowTESTNO))
                        {
                            beforeRow++;
                        }
                        else//如果不等则合并记录下的单元格区域,并记录新的行号和单元格值
                        {
                            if (tempRow < beforeRow)//如果最新行号小于遍历的上一个行号
                            {

                                //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
                                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(tempRow, beforeRow, colList[i], colList[i]));
                                EXCEL.Range rangeProgram = sheet.get_Range(sheet.Cells[tempRow, colList[i]], sheet.Cells[beforeRow, colList[i]]);
                                rangeProgram.Application.DisplayAlerts = false;
                                rangeProgram.Merge(Missing.Value);
                            }
                            tempCellValue = nowCellValue;//更新单元格值                           
                            tempTESTNO = nowTESTNO; //更新检测项值

                            tempRow = j;
                            beforeRow++;
                        }

                    }
                    #endregion
                }
                return;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }

        /// <summary>
        /// 合并指定列,按值相等合并
        /// </summary>
        /// <param name="sheetName">目标sheet</param>
        /// <param name="colList">要合并的单元格所在列</param>
        /// <param name="startRow">开始行</param>
        /// <param name="endRow">结束行</param>
        public void mergeCells(string sourceFile, int sheetIndex, int[] colList, int startRow, int endCol)
        {
            if (File.Exists(sourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(sourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    mergeCells(workBook, sheetIndex, colList, startRow, endCol);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
        }




        /// <summary>
        /// 行合并,按值相等合并
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        private void mergeRowCells(EXCEL.Workbook wb, int sheetIndex, int startRow, int endRow, int endCol)
        {
            try
            {
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

                for (int i = startRow; i <= endRow; i++)//遍历需要合并的行
                {
                    #region 检查数值相等并合并当前列
                    EXCEL.Range tempRow = (EXCEL.Range)sheet.Rows[1];
                    EXCEL.Range tempCell = (EXCEL.Range)sheet.Cells[i, 1];

                    string tempCellValue = getMergerCellValue(sheet, i, 1);
                    int tempCol = 1;//最新列号,作为需要合并的起始列
                    int beforeCol = 1;//之前的列号,作为需要合并的结束列
                    for (int j = 2; j <= endCol; j++)//遍历列的指定行集合
                    {

                        EXCEL.Range tempCell1 = (EXCEL.Range)sheet.Cells[i, j];
                        if (tempCell1 == null) continue;
                        string nowCellValue = getMergerCellValue(sheet, i, j);

                        if (tempCellValue.Equals(nowCellValue))//如果相等则之前的列号+1
                        {
                            beforeCol++;

                        }
                        else//如果不等则合并记录下的单元格区域,并记录新的列号和单元格值
                        {
                            //如果最新列号小于遍历的上一个列号,且单元格值非空
                            if (tempCol < beforeCol && !tempCellValue.Equals(""))
                            {

                                //设置一个合并单元格区域
                                //获取需要合并的单元格的范围
                                EXCEL.Range rangeProgram = sheet.get_Range(sheet.Cells[i, tempCol], sheet.Cells[i, beforeCol]);
                                rangeProgram.Application.DisplayAlerts = false;
                                rangeProgram.Merge(Missing.Value);

                            }
                            tempCellValue = nowCellValue;//更新单元格值
                            tempCol = j;
                            beforeCol++;
                        }
                    }
                    #endregion
                }
                return;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }

        /// <summary>
        /// 行合并,按值相等合并,指定行的列范围
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startRow"></param>
        /// <param name="startCol">起始列索引</param>
        /// <param name="endCol"></param>
        /// <param name="unpivotRange">转置标记列</param>
        /// <param name="unpivotMerge">转置合并标记数组</param>
        private void mergeRows(EXCEL.Workbook wb, int sheetIndex, int startRow, System.Drawing.Point unpivotRange, int endCol, string[] unpivotMerge)
        {
            try
            {
                if (unpivotRange.X < 1 || unpivotRange.Y < 1)
                {
                    return;

                }
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                int endRow = sheet.UsedRange.Rows.Count;

                for (int i = startRow; i <= endRow; i++)//遍历需要合并的行
                {
                    //转置标记列必须为1
                    if ((unpivotMerge.Length > i - startRow) && (unpivotMerge[i - startRow].Equals("1")))
                    {
                        //设置一个合并单元格区域
                        //获取需要合并的单元格的范围
                        EXCEL.Range rangeProgram = sheet.get_Range(sheet.Cells[i, unpivotRange.X], sheet.Cells[i, unpivotRange.Y]);
                        rangeProgram.Application.DisplayAlerts = false;
                        rangeProgram.Merge(Missing.Value);
                    }

                }
                return;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }

        //合并检测项目和分析项
        /// <summary>
        /// 合并检测项目和分析项所在列数据,当值相同时
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startRow"></param>
        /// <param name="maxCol"></param>
        private void mergeRowTestAndAnalyte(EXCEL.Workbook wb, int sheetIndex, int startRow, int maxCol)
        {
            try
            {
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                int endRow = sheet.UsedRange.Rows.Count;

                System.Drawing.Point startColPoint = selectPosition(sheet, "检测项目", endRow, maxCol);
                System.Drawing.Point endColPoint = selectPosition(sheet, "分析项", endRow, maxCol);
                int startCol = startColPoint.Y;
                int endCol = endColPoint.Y;
                //"----以下空白----"的行不用遍历
                for (int i = startRow; i < endRow; i++)//遍历需要合并的行
                {
                    string testValue = getMergerCellValue(sheet, i, startCol);
                    string analyteValue = getMergerCellValue(sheet, i, endCol);
                    if (testValue.Equals(analyteValue))//如果相等则之前的列号+1
                    {
                        //设置一个合并单元格区域
                        //获取需要合并的单元格的范围
                        EXCEL.Range rangeProgram = sheet.get_Range(sheet.Cells[i, startCol], sheet.Cells[i, endCol]);
                        rangeProgram.Application.DisplayAlerts = false;
                        rangeProgram.Merge(Missing.Value);

                    }
                }
                return;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }

        /// <summary>
        /// 合并检测项目和分析项
        /// </summary>
        /// <param name="wb">目标workbook</param>
        /// <param name="sheetIndex"></param>
        /// <param name="str1">检测项目标记字符串</param>
        /// <param name="str2">分析项标记字符串</param>
        private void mergeTestCell(EXCEL.Workbook wb, int sheetIndex, string str1, string str2, int endCol)
        {

            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
            int endRow = sheet.UsedRange.Rows.Count;

            System.Drawing.Point p1 = selectPosition(sheet, str1, endRow, endCol);
            System.Drawing.Point p2 = selectPosition(sheet, str2, endRow, endCol);
            //不在同一行不合并
            if (p1.X != p2.X) return;
            //不是相邻单元格不合并
            if (p1.Y + 1 != p2.Y) return;
            else
            {
                EXCEL.Range rangeProgram = sheet.get_Range(sheet.Cells[p1.X, p1.Y], sheet.Cells[p2.X, p2.Y]);
                rangeProgram.Application.DisplayAlerts = false;
                rangeProgram.Merge(Missing.Value);
            }
            classLims_NPOI.WriteLog("mergeTestCell success", "");
        }

        //处理报告附页的格式调整
        private void reportOneDimDExcelFormat(EXCEL.Workbook wb, int sheetIndex, int[] colseq, int startRow,
            double updHeight, int startCol, int endCol, object[] specialChars, System.Drawing.Point unpivotRange, string[] unpivotMerge)
        {
            ReplaceAll(wb, sheetIndex, specialChars);//替换特殊字符
            mergeRowTestAndAnalyte(wb, sheetIndex, startRow, endCol);//合并检测项目和分析项
            mergeRows(wb, sheetIndex, startRow, unpivotRange, endCol, unpivotMerge); //合并转置列
            mergeCells(wb, sheetIndex, colseq, startRow, endCol);//合并相同检测项目的同列数据
            setAutoRowHeight(wb, sheetIndex, startRow, startCol, endCol, updHeight);//设置行高
            dealMergedAreaInPages_new(wb, sheetIndex, endCol, colseq);//跨页的要分开,只处理需要合并的列,多实测值模板不拆分实测值的合并
            mergeTestCell(wb, sheetIndex, "检测项目", "分析项", endCol);//合并表头的"检测项目","分析项"2个单元格
            stretchLastRowHeight(wb, sheetIndex);//拉伸最后一行,设置边框位置为页底
        }

        //处理报告附页的格式调整
        /// <summary>
        /// 处理报告附页的格式调整
        /// </summary>
        /// <param name="strSourceFile">目标文件</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="colseq">要合并的列索引</param>
        /// <param name="startRow">数据页起始列</param>
        /// <param name="updHeight">调整行高</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="specialChars">替换字符数组</param>
        /// <param name="unpivotSeq">转置列列头</param>
        /// <param name="unpivotRange">转置标记列</param>
        /// <param name="unpivotMerge">转置合并标记数组</param>
        /// <returns></returns>
        public bool reportOneDimDExcelFormat(string strSourceFile, int sheetIndex, int[] colseq, int startRow,
            double updHeight, int startCol, int endCol, object[] specialChars, System.Drawing.Point unpivotRange, string[] unpivotMerge)
        {
            bool flag = true;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    //先使用分页视图打开,EXCEl获取 HPageBreaks 需要在分页视图中
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlPageBreakPreview;
                    reportOneDimDExcelFormat(workBook, sheetIndex, colseq, startRow,
                        updHeight, startCol, endCol, specialChars, unpivotRange, unpivotMerge);
                    //再还原为普通视图
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlNormalView;
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return flag;
        }

        //使用Excel的全部替换功能,替换特殊字符
        private static void ReplaceAll(EXCEL.Workbook wb, int sheetIndex, object[] dArray)
        {
            object missing = System.Reflection.Missing.Value;
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

            EXCEL.Range excelRange = (EXCEL.Range)sheet.Rows;
            string oldStr;
            string newStr;
            Dictionary<string, string> hashMap = new Dictionary<string, string>();
            hashMap = classLims_NPOI.dArray2Dictionary(dArray);
            foreach (var oneMapPoint in hashMap)
            {
                string key = oneMapPoint.Key.ToString();
                string value = oneMapPoint.Value.ToString();
                oldStr = key;
                newStr = value;
                //xlPart代表匹配任一部分搜索文本。, xlWhole代表匹配全部搜索文本。
                excelRange.Replace(oldStr, newStr, EXCEL.XlLookAt.xlPart, EXCEL.XlSearchOrder.xlByRows, missing, missing, missing, missing);
            }

            excelRange = null;

        }

        //使用Excel的全部替换功能,替换特殊字符
        public static bool ReplaceAll(string strSourceFile, int sheetIndex, object[] dArray)
        {
            bool flag = true;
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    ReplaceAll(workBook, sheetIndex, dArray);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return flag;

        }

        /// <summary>
        /// 设置自动行高,可处理合并后的行高
        /// </summary>
        /// <param name="wb">工作簿</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="startRow">起始行</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="updHeight">调整行高,根号倍数放大行高, 20代表增加20%倍根号(原行高)</param>
        public static void setAutoRowHeight(EXCEL.Workbook wb, int sheetIndex, int startRow, int startCol, int endCol, double updHeight)
        {
            try
            {
                //wb.Application.Run("AutoHeight");
                object missing = System.Reflection.Missing.Value;
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                int endRow = sheet.UsedRange.Rows.Count;

                double autoRowheight = 0;
                //设置行高时,不应该修改最后一行的"----以下空白-----"行高
                for (int i = startRow; i < endRow; i++)
                {
                    //classLims_NPOI.WriteLog("rowIndex:" + i, "");
                    EXCEL.Range row = (EXCEL.Range)sheet.Rows[i];
                    autoRowheight = getRowHeight(wb, sheetIndex, i, startCol, endCol);
                    //classLims_NPOI.WriteLog("获取行高:" + autoRowheight + "; 偏移行高:" + updHeight, "");
                    if (autoRowheight > 0)
                    {
                        //根号倍数放大行高
                        double sumRowHeight = autoRowheight + (updHeight * Math.Sqrt(autoRowheight) / 100);
                        if (sumRowHeight > 409) sumRowHeight = 409;
                        //double sumRowHeight = autoRowheight + updHeight;
                        row.RowHeight = sumRowHeight;
                        //classLims_NPOI.WriteLog("success:" + i, "");
                    }

                }
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
        }

        /// <summary>
        /// 设置自动行高,可处理合并后的行高
        /// </summary>
        /// <param name="strSourceFile">要修改的excel</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="startRow">起始行</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="updHeight">调整行高</param>
        public static void setAutoRowHeight(string strSourceFile, int sheetIndex, int startRow, int startCol, int endCol, double updHeight)
        {
            if (File.Exists(strSourceFile))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    excel.DisplayAlerts = false;
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    setAutoRowHeight(workBook, sheetIndex, startRow, startCol, endCol, updHeight);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }
            }
            return;
        }


        /// <summary>
        /// 查询值在sheet的位置,X:行号,Y:列号
        /// </summary>
        /// <param name="sheet">工作表名</param>
        /// <param name="value">标志字符串</param>
        /// <returns>Point对象,X:行号,Y:列号</returns>
        public static System.Drawing.Point selectPosition(EXCEL.Worksheet sheet, string value, int endRow, int endCol)
        {
            System.Drawing.Point p = new System.Drawing.Point();
            p.X = -1;
            p.Y = -1;
            try
            {
                int minRow = 1;
                int maxRow = endRow;
                for (int i = minRow; i <= maxRow; i++)
                {
                    EXCEL.Range row = (EXCEL.Range)sheet.Rows[i];
                    if (row == null)
                    {
                        continue;
                    }
                    for (int j = 1; j <= endCol; j++)
                    {
                        string cellValue = getMergerCellValue(sheet, i, j); ;
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
                classLims_NPOI.WriteLog(ex, "");
                return p;
            }
        }

        /// <summary>
        /// 拉伸表格到A4大小
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public static void stretchLastRowHeight(string filePath, int sheetIndex)
        {
            if (File.Exists(filePath))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbooks workBooks = null;
                EXCEL.Workbook workBook = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(filePath, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    stretchLastRowHeight(workBook, sheetIndex);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
            }
        }

        /// <summary>
        /// 拉伸表格到A4大小
        /// </summary>
        /// <param name="Workbook">excel Workbook对象</param>
        /// <param name="sheetIndex"></param>
        /// <returns>是否成功</returns>
        public static void stretchLastRowHeight(EXCEL.Workbook wb, int sheetIndex)
        {
            try
            {
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
                int endCol = sheet.UsedRange.Columns.Count;
                int startHeadRow, endHeadRow;
                int[] iRange = getRepeatingRowsRange(sheet, endCol);
                if (iRange == null)
                {
                    startHeadRow = 0;
                    endHeadRow = 0;
                }
                else
                {
                    startHeadRow = iRange[0];
                    endHeadRow = iRange[1];
                }
                List<int> arrRf = getNewPageFirstRow(sheet);
                int fRow;
                if (arrRf == null || arrRf.Count == 0)
                {
                    fRow = endHeadRow + 1;
                }
                else
                {
                    fRow = arrRf[arrRf.Count - 1];
                }
                double totalH = 0; //总行高
                                   //高度换算 1厘米＝28.346456692913389磅
                double headH = 0; //表头高度
                headH = sheet.PageSetup.TopMargin + sheet.PageSetup.BottomMargin;
                if (iRange == null)
                {
                    headH += 0;
                }
                else
                {
                    headH += (double)sheet.get_Range(sheet.Cells[iRange[0], iRange[2]], sheet.Cells[iRange[1], iRange[3]]).Height;
                }

                for (int i = fRow; i <= 65535; i++)
                {
                    EXCEL.Range row = (EXCEL.Range)sheet.Rows[i];
                    if (row == null) continue;

                    double tempH;
                    if ((bool)((EXCEL.Range)sheet.Rows[i]).Hidden == true)
                    {
                        tempH = 0;
                    }
                    else
                    {
                        tempH = (double)((EXCEL.Range)sheet.Rows[i]).Height;
                    }

                    if (System.Convert.ToInt32(totalH + tempH) >= System.Convert.ToInt32(PAGE_HEIGHT - (1 * headH)))//超过一页
                    {
                        //第一次找到的页尾,就是最后一页的页尾

                        //先合并"----以下空白----"行
                        int lRow = sheet.UsedRange.Rows.Count;
                        EXCEL.Range lastMarkRange = sheet.get_Range(sheet.Cells[lRow, 1], sheet.Cells[lRow, endCol]);
                        lastMarkRange.Merge();
                        //垂直居上,水平居中
                        lastMarkRange.VerticalAlignment = XlVAlign.xlVAlignTop;
                        lastMarkRange.HorizontalAlignment = XlVAlign.xlVAlignCenter;

                        EXCEL.Range lastRange = sheet.get_Range(sheet.Cells[lRow, 1], sheet.Cells[i - 1, endCol]);
                        //设置边框为外框线
                        object missing = System.Reflection.Missing.Value;
                        lastRange.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
                        lastRange.BorderAround2(XlLineStyle.xlContinuous, EXCEL.XlBorderWeight.xlThin, EXCEL.XlColorIndex.xlColorIndexAutomatic, missing, missing);
                        break;
                    }
                    else
                    {
                        totalH = totalH + tempH;
                    }
                }
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
        }

        public void protectWorkBook(string filePath, string psw)
        {
            if (File.Exists(filePath))
            {
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbooks workBooks = null;
                EXCEL.Workbook workBook = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(filePath, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    /*参数
                     * Password
                     * 工作簿的密码，区分大小写。如果省略此参数，则无需使用密码即可取消保护工作簿。否则，必须指定密码才能取消保护工作簿。
                     * Structure
                     * 如果为 true，则保护工作簿的结构（工作表的相对位置）。默认值为 false。
                     * Windows
                     * 如果为 true，则保护工作簿窗口。如果省略此参数，则窗口不受保护。2007之后的版本已经取消"保护窗口"这个功能了
                     */
                    workBook.Protect(psw, true, missing);
                    workBook.Save();
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                }
                finally
                {
                    if (workBook != null)
                    {
                        workBook.Close(false, missing, missing);
                        Marshal.ReleaseComObject(workBook);
                        //Marshal.FinalReleaseComObject(workBook);
                        workBook = null;
                    }
                    if (workBooks != null)
                    {
                        workBooks.Close();
                        Marshal.ReleaseComObject(workBooks);
                        workBook = null;
                    }
                    if (excel != null)
                    {
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        //Marshal.FinalReleaseComObject(excel);
                        excel = null;

                        //flag = KillSpecialExcel(excel);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
            }
        }

        //保护工作表
        /// <summary>
        /// 保护工作表
        /// </summary>
        /// <param name="wb">工作簿</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="password">保护密码</param>
        /// 
        /// <param name="DrawingObjects">如果为 True，则保护形状。默认值是 True。对应excel中"编辑对象"</param>
        /// <param name="Contents">如果为 True，则保护内容。对于图表，这样会保护整个图表。对于工作表，这样会保护锁定的单元格。默认值是 True。</param>
        /// <param name="Scenarios">如果为 True，则保护方案。此参数仅对工作表有效。默认值是 True。</param>
        /// 
        /// <param name="UserInterfaceOnly">如果为 True，则保护用户界面，但不保护宏。如果省略此参数，则既保护宏也保护用户界面。</param>
        /// <param name="AllowFormattingCells">如果为 True，则允许用户为受保护的工作表上的任意单元格设置格式。默认值是 False。</param>
        /// <param name="AllowFormattingColumns">如果为 True，则允许用户为受保护的工作表上的任意列设置格式。默认值是 False。</param>
        /// 
        /// <param name="AllowFormattingRows">如果为 True，则允许用户为受保护的工作表上的任意行设置格式。默认值是 False。</param>
        /// <param name="AllowInsertingColumns">如果为 True，则允许用户在受保护的工作表上插入列。默认值是 False。</param>
        /// <param name="AllowInsertingRows">如果为 True，则允许用户在受保护的工作表上插入行。默认值是 False。</param>
        /// 
        /// <param name="AllowInsertingHyperlinks">如果为 True，则允许用户在受保护的工作表中插入超链接。默认值是 False。</param>
        /// <param name="AllowDeletingColumns">如果为 True，则允许用户在受保护的工作表上删除列，要删除的列中的每个单元格都被解除锁定。默认值是 False。</param>
        /// <param name="AllowDeletingRows">如果为 True，则允许用户在受保护的工作表上删除行，要删除的行中的每个单元格都被解除锁定。默认值是 False。</param>
        /// 
        /// <param name="AllowSorting">如果为 True，则允许用户在受保护的工作表上进行排序。排序区域中的每个单元格必须是解除锁定的或取消保护的。默认值是 False。</param>
        /// <param name="AllowFiltering">如果为 True，则允许用户在受保护的工作表上设置筛选。用户可以更改筛选条件，但是不能启用或禁用自动筛选功能。用户也可以在已有的自动筛选功能上设置筛选。默认值是 False。</param>
        /// <param name="AllowUsingPivotTables">如果为 True，则允许用户在受保护的工作表上使用数据透视表。默认值是 False。</param>
        public void protectWorkSheet(EXCEL.Workbook wb, int sheetIndex, string password,
            object DrawingObjects, object Contents, object Scenarios,
            object UserInterfaceOnly, object AllowFormattingCells, object AllowFormattingColumns,
            object AllowFormattingRows, object AllowInsertingColumns, object AllowInsertingRows,
            object AllowInsertingHyperlinks, object AllowDeletingColumns, object AllowDeletingRows,
            object AllowSorting, object AllowFiltering, object AllowUsingPivotTables)
        {
            object missing = Missing.Value;
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];
            sheet.Protect(password,
                DrawingObjects, Contents, Scenarios,
                UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns,
                AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows,
                AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows,
                AllowSorting, AllowFiltering, AllowUsingPivotTables);
        }

        //保护文件下指定工作表
        public void protectWorkSheet(string path, int sheetIndex, string password,
            object DrawingObjects, object Contents, object Scenarios,
            object UserInterfaceOnly, object AllowFormattingCells, object AllowFormattingColumns,
            object AllowFormattingRows, object AllowInsertingColumns, object AllowInsertingRows,
            object AllowInsertingHyperlinks, object AllowDeletingColumns, object AllowDeletingRows,
            object AllowSorting, object AllowFiltering, object AllowUsingPivotTables)
        {
            object missing = Missing.Value;
            EXCEL.ApplicationClass excel = null;
            EXCEL.Workbook wb = null;
            EXCEL.Workbooks workBooks = null;
            try
            {
                excel = new EXCEL.ApplicationClass();
                excel.DisplayAlerts = false;
                workBooks = excel.Workbooks;
                wb = workBooks.Open(path, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing);

                protectWorkSheet(wb, sheetIndex, password,
                 DrawingObjects, Contents, Scenarios,
                 UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns,
                 AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows,
                 AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows,
                 AllowSorting, AllowFiltering, AllowUsingPivotTables);

                wb.Save();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
            finally
            {
                if (wb != null)
                {
                    //wb.Close(false, missing, false);
                    wb.Close(false, missing, missing);
                    int i = Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (workBooks != null)
                {
                    workBooks.Close();
                    int i = Marshal.ReleaseComObject(workBooks);
                    workBooks = null;
                }
                if (excel != null)
                {
                    excel.Quit();
                    int i = Marshal.ReleaseComObject(excel);
                    excel = null;
                }

            }
            
        }

        //使用OFFICE组件设置行高加固定值,修改单位为磅
        /// <summary>
        /// 使用OFFICE组件设置行高加固定值,修改单位为磅
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="updHeight"></param>
        private static void updateSheesRowHeight(EXCEL.Workbook wb, int sheetIndex, int startRow, int endRow, double updHeight)
        {
            object missing = System.Reflection.Missing.Value;
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex];

            for (int i = startRow; i <= endRow; i++)
            {
                EXCEL.Range excelRange = (EXCEL.Range)sheet.Rows[i, missing];
                excelRange.RowHeight = (double)excelRange.RowHeight + updHeight;
                excelRange = null;
            }
        }


    }
}
