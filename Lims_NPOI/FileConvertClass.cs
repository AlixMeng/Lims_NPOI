using System;
using System.IO;
using System.Runtime.InteropServices;
using EXCEL = Microsoft.Office.Interop.Excel;
using WORD = Microsoft.Office.Interop.Word;
using System.Reflection;
using novapiLib80;
using System.Collections.Generic;
//using PDFMAKERAPILib;

namespace nsLims_NPOI
{
    #region  文件转换类
    /// <summary>
    /// 文件转换类,使用控件或另存为功能
    /// </summary>
    class FileConvertClass
    {

        public void novaFunc()
        {
            NovaPdfOptions80Class np8 = new NovaPdfOptions80Class();
        }

        /// <summary>
        /// 报告批准时,excel转pdf,封面和首页只取第一页
        /// </summary>
        /// <param name="workBook">工作簿对象</param>
        /// <param name="toPath">输出路径</param>
        /// <returns>成功或失败</returns>
        public static bool ApproveExcelToPdf(EXCEL.Workbook workBook, string toPath)
        {
            bool flag = false;
            object missing = Type.Missing;
            try
            {
                if (toPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的目标文件路径不能为空。");
                }

                List<string> listSheetPdf = new List<string>();
                string toSheetPdf = "";
                for (int i = 1; i <= workBook.Sheets.Count; i++)//循环遍历sheet,
                {
                    toSheetPdf = "C:\\Windowd\\Temp\\" + Guid.NewGuid().ToString()+".pdf";
                    SaveExcelWorkSheetAsPDFWithPage(workBook, i, toSheetPdf, 1);//转化sheet页为pdf,只转一页
                    listSheetPdf.Add(toSheetPdf);
                }
                MergePDF mp = new MergePDF();
                mp.Merge(listSheetPdf.ToString(), toPath);
                flag = true;
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                flag = false;
            }
            return flag;
        }


        #region 把指定EXCEL转换成PDF
        /// <summary>
        /// EXCEL转PDF
        /// </summary>
        /// <param name="strSourceFile">要转换的EXCEL文件路径</param>
        /// <param name="strTargetFile">目标文件</param>
        /// <returns>转换结果：TRUE FALSE</returns>
        public bool ExcelConvertTOPDF(string strSourceFile, string strTargetFile)
        {
            bool flag = false;
            if (File.Exists(strSourceFile))
            {
                //转换成的上档格式PDF
                EXCEL.XlFixedFormatType targetType = EXCEL.XlFixedFormatType.xlTypePDF;
                object targetFile = strTargetFile;
                object missing = Type.Missing;
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                EXCEL.Worksheet sheet = null;
                try
                {
                    excel = new EXCEL.ApplicationClass();
                    //excel.DisplayAlerts = true;
                    workBooks = excel.Workbooks;
                    workBook = workBooks.Open(strSourceFile, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);

                    //设置格式，导出成PDF
                    sheet = (EXCEL.Worksheet)workBook.Worksheets[1];//下载从1开始

                    //把sheet设置成横向
                    //sheet.PageSetup.Orientation = EXCEL.XlPageOrientation.xlLandscape;
                    //可以设置sheet页的其他相关设置，不列举
                    sheet.ExportAsFixedFormat(targetType, targetFile, EXCEL.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                    flag = true;
                }
                catch (Exception ex)
                {
                    classLims_NPOI.WriteLog(ex, "");
                    flag = false;
                }
                finally
                {
                    if (sheet != null)
                    {

                        Marshal.ReleaseComObject(sheet);
                        //Marshal.FinalReleaseComObject(sheet);
                        sheet = null;
                    }
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
        /// 打印整个workbook,使用另存为功能转pdf
        /// </summary>
        /// <param name="fromExcelPath">源excel路径</param>
        /// <param name="toPath">输出路径</param>
        /// <returns>成功或失败</returns>
        public bool SaveExcelWorkbookAsPDF(string fromExcelPath, string toPath)
        {
            bool flag = false;
            if (File.Exists(fromExcelPath))
            {
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                object missing = Type.Missing;
                try
                {
                    if (fromExcelPath.Length == 0)
                    {
                        flag = false;
                        throw new Exception("需要转换的源文件路径不能为空。");
                    }
                    if (toPath.Length == 0)
                    {
                        flag = false;
                        throw new Exception("需要转换的目标文件路径不能为空。");
                    }

                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    Type type = workBooks.GetType();
                    workBook = workBooks.Open(fromExcelPath, missing, true,
                            missing, missing, missing, missing, missing,
                            missing, missing, missing, missing, missing,
                            missing, missing);

                    //先使用分页视图打开,EXCEl获取 HPageBreaks 需要在分页视图中
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlPageBreakPreview;
                    //int hpbCount = classExcelMthd.getSheetPageCount(workBook, 1);
                    //按照设置好的打印区域发布为pdf
                    workBook.ExportAsFixedFormat(
                        EXCEL.XlFixedFormatType.xlTypePDF,
                        toPath,
                        EXCEL.XlFixedFormatQuality.xlQualityStandard,//可设置为 xlQualityStandard 或 xlQualityMinimum。
                        true,//包含文档属性
                        false, //如果设置为 True，则忽略在发布时设置的任何打印区域。如果设置为 False，则使用在发布时设置的打印区域。
                        Type.Missing,//发布的起始页码。如果省略此参数，则从起始位置开始发布。
                        Type.Missing,//发布的终止页码。如果省略此参数，则发布至最后一页。
                        false, //是否发布文件后在查看器中显示文件。
                        Type.Missing);
                    //再还原为普通视图
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlNormalView;
                    flag = true;
                }
                catch (Exception exception)
                {
                    classLims_NPOI.WriteLog(exception, "");
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
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return flag;

        }

        /// <summary>
        /// 打印workbook,使用另存为功能转pdf,指定打印页码
        /// </summary>
        /// <param name="fromExcelPath">源excel路径</param>
        /// <param name="toPath">输出路径</param>
        /// <param name="pageFrom">打印起始页码</param>
        /// <param name="pageTo">打印结束页码</param>
        /// <returns>成功或失败</returns>
        public bool SaveExcelWorkbookAsPDFWithPage(string fromExcelPath, string toPath, int pageFrom, int pageTo)
        {
            bool flag = false;
            if (File.Exists(fromExcelPath))
            {
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                object missing = Type.Missing;
                try
                {
                    if (fromExcelPath.Length == 0)
                    {
                        flag = false;
                        throw new Exception("需要转换的源文件路径不能为空。");
                    }
                    if (toPath.Length == 0)
                    {
                        flag = false;
                        throw new Exception("需要转换的目标文件路径不能为空。");
                    }

                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    Type type = workBooks.GetType();
                    workBook = workBooks.Open(fromExcelPath, missing, true,
                            missing, missing, missing, missing, missing,
                            missing, missing, missing, missing, missing,
                            missing, missing);

                    //先使用分页视图打开,EXCEl获取 HPageBreaks 需要在分页视图中
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlPageBreakPreview;
                    //int hpbCount = classExcelMthd.getSheetPageCount(workBook, 1);
                    //按照设置好的打印区域发布为pdf
                    workBook.ExportAsFixedFormat(
                        EXCEL.XlFixedFormatType.xlTypePDF,
                        toPath,
                        EXCEL.XlFixedFormatQuality.xlQualityStandard,//可设置为 xlQualityStandard 或 xlQualityMinimum。
                        true,//包含文档属性
                        false, //如果设置为 True，则忽略在发布时设置的任何打印区域。如果设置为 False，则使用在发布时设置的打印区域。
                        pageFrom,//发布的起始页码。如果省略此参数，则从起始位置开始发布。
                        pageTo,//发布的终止页码。如果省略此参数，则发布至最后一页。
                        false, //是否发布文件后在查看器中显示文件。
                        Type.Missing);
                    //再还原为普通视图
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlNormalView;
                    flag = true;
                }
                catch (Exception exception)
                {
                    classLims_NPOI.WriteLog(exception, "");
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
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return flag;

        }

        /// <summary>
        /// 打印workbook,使用另存为功能转pdf,指定打印页码
        /// </summary>
        /// <param name="workBook">工作簿对象</param>
        /// <param name="toPath">输出路径</param>
        /// <param name="pageFrom">打印起始页码</param>
        /// <param name="pageTo">打印结束页码</param>
        /// <returns>成功或失败</returns>
        public static bool SaveExcelWorkbookAsPDFWithPage(EXCEL.Workbook workBook, string toPath, int pageFrom, int pageTo)
        {
            bool flag = false;
            object missing = Type.Missing;
            try
            {
                if (toPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的目标文件路径不能为空。");
                }

                //按照设置好的打印区域发布为pdf
                workBook.ExportAsFixedFormat(
                    EXCEL.XlFixedFormatType.xlTypePDF,
                    toPath,
                    EXCEL.XlFixedFormatQuality.xlQualityStandard,//可设置为 xlQualityStandard 或 xlQualityMinimum。
                    true,//包含文档属性
                    false, //如果设置为 True，则忽略在发布时设置的任何打印区域。如果设置为 False，则使用在发布时设置的打印区域。
                    pageFrom,//发布的起始页码。如果省略此参数，则从起始位置开始发布。
                    pageTo,//发布的终止页码。如果省略此参数，则发布至最后一页。
                    false, //是否发布文件后在查看器中显示文件。
                    Type.Missing);
                //再还原为普通视图
                flag = true;
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                flag = false;
            }
            return flag;
        }

        /// <summary>
        /// 保存sheet页为pdf,指定转化页数
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="toPath"></param>
        /// <param name="toPage"></param>
        /// <returns></returns>
        public static bool SaveExcelWorkSheetAsPDFWithPage(EXCEL.Workbook workBook, int sheetIndex, string toPath, int toPage)
        {
            bool flag = false;
            object missing = Type.Missing;
            try
            {
                if (toPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的目标文件路径不能为空。");
                }

                EXCEL.Worksheet sheet = (EXCEL.Worksheet)workBook.Sheets[sheetIndex];
                //sheet.Activate();
                sheet.ExportAsFixedFormat(
                    EXCEL.XlFixedFormatType.xlTypePDF,
                    toPath,
                    EXCEL.XlFixedFormatQuality.xlQualityStandard,//可设置为 xlQualityStandard 或 xlQualityMinimum。
                    true,//包含文档属性
                    false, //如果设置为 True，则忽略在发布时设置的任何打印区域。如果设置为 False，则使用在发布时设置的打印区域。
                    1,//发布的起始页码。如果省略此参数，则从起始位置开始发布。
                    toPage,//发布的终止页码。如果省略此参数，则发布至最后一页。
                    false, //是否发布文件后在查看器中显示文件。
                    Type.Missing);
                flag = true;
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                flag = false;
            }
            return flag;
        }
        
        /// <summary>
        /// excel批量转为pdf
        /// </summary>
        /// <param name="fromExcellist"></param>
        /// <param name="printDir"></param>
        /// <returns></returns>
        public string[] ExcellistToPDF(string fromExcellist, string printDir)
        {
            string[] strArrays = fromExcellist.Split(new char[] { ',' });//逗号分隔数组
            List<string> aPdfFile = new List<string>();

            object missing = Type.Missing;
            EXCEL.ApplicationClass excel = null;
            EXCEL.Workbooks workBooks = null;
            EXCEL.Workbook wb = null;
            try
            {

                excel = new EXCEL.ApplicationClass();
                excel.DisplayAlerts = false;
                workBooks = excel.Workbooks;
                string sPdfFile = "";
                if (printDir == "") printDir = "C:\\Windows\\Temp\\";

                for (int i = 0; i < strArrays.Length; i++)
                {
                    wb = workBooks.Open(strArrays[i], missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing,
                        missing, missing);
                    sPdfFile = printDir + System.Guid.NewGuid().ToString() + ".pdf";
                    wb.ExportAsFixedFormat(
                            EXCEL.XlFixedFormatType.xlTypePDF,
                            sPdfFile,
                            EXCEL.XlFixedFormatQuality.xlQualityStandard,//可设置为 xlQualityStandard 或 xlQualityMinimum。
                            true,//包含文档属性
                            false, //如果设置为 True，则忽略在发布时设置的任何打印区域。如果设置为 False，则使用在发布时设置的打印区域。
                            Type.Missing,//发布的起始页码。如果省略此参数，则从起始位置开始发布。
                            Type.Missing,//发布的终止页码。如果省略此参数，则发布至最后一页。
                            false, //是否发布文件后在查看器中显示文件。
                            Type.Missing);
                    wb.Save();
                    if (wb != null)
                    {
                        wb.Close(false, missing, missing);
                        Marshal.ReleaseComObject(wb);
                        wb = null;
                    }
                    aPdfFile.Add(sPdfFile);
                }
                return aPdfFile.ToArray();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return null;
            }
            finally
            {
                if (workBooks != null)
                {
                    workBooks.Close();
                    Marshal.ReleaseComObject(workBooks);
                    workBooks = null;
                }
                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    excel = null;

                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public bool ExcelWorkbookPrintToPDF(string fromExcelPath, string toPath)
        {
            bool flag = false;
            if (File.Exists(fromExcelPath))
            {
                EXCEL.ApplicationClass excel = null;
                EXCEL.Workbook workBook = null;
                EXCEL.Workbooks workBooks = null;
                object missing = Type.Missing;
                try
                {
                    if (fromExcelPath.Length == 0)
                    {
                        flag = false;
                        throw new Exception("需要转换的源文件路径不能为空。");
                    }
                    if (toPath.Length == 0)
                    {
                        flag = false;
                        throw new Exception("需要转换的目标文件路径不能为空。");
                    }

                    excel = new EXCEL.ApplicationClass();
                    workBooks = excel.Workbooks;
                    Type type = workBooks.GetType();
                    workBook = workBooks.Open(fromExcelPath, missing, true,
                            missing, missing, missing, missing, missing,
                            missing, missing, missing, missing, missing,
                            missing, missing);

                    //先使用分页视图打开,EXCEl获取 HPageBreaks 需要在分页视图中
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlPageBreakPreview;
                    //int hpbCount = classExcelMthd.getSheetPageCount(workBook, 1);
                    //按照设置好的打印区域发布为pdf
                    workBook.PrintOutEx(missing, missing, missing, false, missing,
                        true, false, "ZZY", true);
                    //再还原为普通视图
                    excel.ActiveWindow.View = EXCEL.XlWindowView.xlNormalView;
                    flag = true;
                }
                catch (Exception exception)
                {
                    classLims_NPOI.WriteLog(exception, "");
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
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return flag;

        }

        #endregion

        #region 结束EXCEL.EXE进程的方法
        /// <summary>
        /// 结束EXCEL.EXE进程的方法
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lpdwProcessId"></param>
        /// <returns></returns>
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        public bool KillSpecialExcel(Microsoft.Office.Interop.Excel.Application m_objExcel)
        {
            try
            {
                if (m_objExcel != null)
                {
                    int lpdwProcessId;
                    //获取进程的PID
                    GetWindowThreadProcessId(new IntPtr(m_objExcel.Hwnd), out lpdwProcessId);

                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
                return true;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return false;
            }
        }
        #endregion

        #region 把指定WORD转换成PDF

        /// <summary>
        /// 把指定WORD转换成PDF 
        /// </summary>
        /// <param name="strSourceFile">要转换的Word文档</param>
        /// <param name="strTargetFile">转换成的结果文件</param>
        /// <returns></returns>
        public bool WordConvertTOPDF(object strSourceFile, object strTargetFile)
        {
            try
            {
                bool flag = true;
                if (File.Exists(strSourceFile.ToString()))
                {
                    object Nothing = System.Reflection.Missing.Value;
                    //创建一个名为WordApp的组件对象 
                    WORD.Application wordApp = new WORD.ApplicationClass();
                    //创建一个名为WordDoc的文档对象并打开
                    WORD.Document doc = wordApp.Documents.Open(ref strSourceFile, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);


                    //设置保存的格式 
                    object filefarmat = WORD.WdSaveFormat.wdFormatPDF;
                    //保存为PDF
                    doc.SaveAs(ref strTargetFile, ref filefarmat, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    //关闭文档对象
                    object saveOption = WORD.WdSaveOptions.wdDoNotSaveChanges;
                    ((Microsoft.Office.Interop.Word._Document)doc).Close(ref saveOption, ref Nothing, ref Nothing);
                    //推出组建   
                    wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
                }
                return flag;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return false;
            }


        }

        /// <summary>
        /// 打印整个word,使用另存为功能转pdf
        /// </summary>
        /// <param name="fromWordPath"></param>
        /// <param name="toPath"></param>
        /// <returns>成功或失败</returns>
        public bool SaveWordAsPDF(string fromWordPath, string toPath)
        {
            bool flag = true;
            WORD.ApplicationClass applicationClass = null;
            WORD.Document doc = null;
            try
            {
                if (fromWordPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的源文件路径不能为空。");
                }
                if (!File.Exists(fromWordPath))
                {
                    flag = false;
                    throw new Exception("需要转换的源文件不存在。");
                }
                if (toPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的目标文件路径不能为空。");
                }
                applicationClass = new WORD.ApplicationClass();
                applicationClass.GetType();
                object obj = fromWordPath;
                object[] objArray = new object[] { obj, true, true };
                object oMissing = Missing.Value;
                object oTrue = true;
                object oFalse = false;
                object Copies = 1; //打印份数
                object wdPrintFrom = 1;//打印的起始页码
                object wdPrintTo = 1;//打印的结束页码
                object doNotSaveChanges = WORD.WdSaveOptions.wdDoNotSaveChanges;
                //打开要打印的文件
                doc = applicationClass.Documents.Open(
                    fromWordPath,
                    ref oFalse, //如果该属性为 True，则当文件不是 Microsoft Word 格式时，将显示“转换文件”对话框。
                    ref oTrue, //如果该属性值为 True，则以只读方式打开文档。
                    ref oFalse,//如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中
                    ref oMissing,
                    ref oMissing,
                    ref oFalse, //控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。
                    ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);
               
                
                doc.ExportAsFixedFormat(
                    toPath,
                    WORD.WdExportFormat.wdExportFormatPDF,
                    false,//设为转换后不打开
                    WORD.WdExportOptimizeFor.wdExportOptimizeForPrint,//指定进行屏幕优化还是打印优化 wdExportOptimizeForPrint、wdExportOptimizeForOnScreen。
                    WORD.WdExportRange.wdExportAllDocument, 0, 0,//转换页设为 全部
                    WORD.WdExportItem.wdExportDocumentContent,//指定导出过程是仅包括文本，还是同时包括文本和标记。此处代表仅包括文本
                    true,// 如果要在新文件中包含文档属性，则为 true；否则为 false。
                    true,// 如果要在源文档具有信息权限管理 (IRM) 保护时将 IRM 权限复制到 XPS 文档，则为 true；否则为 false。 默认值为 true。 
                    WORD.WdExportCreateBookmarks.wdExportCreateWordBookmarks,
                    true,// 如果要包含额外数据（如有关内容的流和逻辑组织的信息）来协助使用屏幕读取器，则为 true；否则为 false。 默认值为 true。 
                    true,// 如果要包含文本的位图，则为 true；如果要引用文本字体，则为 false。 如果字体许可证不允许在 PDF 文件中嵌入某种字体，则将此参数设置为 true。 如果将此参数设置为 false，则当指定字体不可用时，查看者的计算机会替换合适的字体。 默认值为 true。 
                    false,//    如果要将 PDF 使用范围限制为按照 ISO 19005-1 进行标准化的 PDF 子集，则为 true；否则为 false。如果将此参数设置为 true，则生成的文件会是更加可靠的独立文件，但这些文件可能会更大，或者由于格式限制而显示更多的视觉瑕疵。默认值为 false。
                    oMissing);

                flag = true;
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                flag = false;
            }
            finally
            {
                if (doc != null)
                {
                    //关闭WORD文件
                    ((WORD._Document)doc).Close(WORD.WdSaveOptions.wdDoNotSaveChanges, Missing.Value, Missing.Value);

                    doc = null;
                }
                if (applicationClass != null)
                {
                    //退出WORD程序
                    ((WORD._Application)applicationClass).Quit(Missing.Value, Missing.Value, Missing.Value);
                    applicationClass = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return flag;
        }

        //word批量转pdf
        public string[] WordlistToPDF(string fromWordlist, string printDir)
        {
            WORD.ApplicationClass applicationClass = null;
            WORD.Document doc = null;
            try
            {
                string[] strArrays = fromWordlist.Split(new char[] { ',' });//逗号分隔数组
                List<string> aPdfFile = new List<string>();
                string sPdfFile = "";
                if (printDir == "") printDir = "C:\\Windows\\Temp\\";

                applicationClass = new WORD.ApplicationClass();
                applicationClass.GetType();
                object oMissing = Missing.Value;
                object oTrue = true;
                object oFalse = false;
                object Copies = 1; //打印份数
                object wdPrintFrom = 1;//打印的起始页码
                object wdPrintTo = 1;//打印的结束页码
                object doNotSaveChanges = WORD.WdSaveOptions.wdDoNotSaveChanges;
                for (int i = 0; i < strArrays.Length; i++)
                {
                    //打开要打印的文件
                    doc = applicationClass.Documents.Open(
                        strArrays[i],
                        ref oFalse, //如果该属性为 True，则当文件不是 Microsoft Word 格式时，将显示“转换文件”对话框。
                        ref oTrue, //如果该属性值为 True，则以只读方式打开文档。
                        ref oFalse,//如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中
                        ref oMissing,
                        ref oMissing,
                        ref oFalse, //控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。
                        ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);

                    sPdfFile = printDir + System.Guid.NewGuid().ToString() + ".pdf";
                    doc.ExportAsFixedFormat(
                        sPdfFile,
                        WORD.WdExportFormat.wdExportFormatPDF,
                        false,//设为转换后不打开
                        WORD.WdExportOptimizeFor.wdExportOptimizeForPrint,//指定进行屏幕优化还是打印优化 wdExportOptimizeForPrint、wdExportOptimizeForOnScreen。
                        WORD.WdExportRange.wdExportAllDocument, 0, 0,//转换页设为 全部
                        WORD.WdExportItem.wdExportDocumentContent,//指定导出过程是仅包括文本，还是同时包括文本和标记。此处代表仅包括文本
                        true,// 如果要在新文件中包含文档属性，则为 true；否则为 false。
                        true,// 如果要在源文档具有信息权限管理 (IRM) 保护时将 IRM 权限复制到 XPS 文档，则为 true；否则为 false。 默认值为 true。 
                        WORD.WdExportCreateBookmarks.wdExportCreateWordBookmarks,
                        true,// 如果要包含额外数据（如有关内容的流和逻辑组织的信息）来协助使用屏幕读取器，则为 true；否则为 false。 默认值为 true。 
                        true,// 如果要包含文本的位图，则为 true；如果要引用文本字体，则为 false。 如果字体许可证不允许在 PDF 文件中嵌入某种字体，则将此参数设置为 true。 如果将此参数设置为 false，则当指定字体不可用时，查看者的计算机会替换合适的字体。 默认值为 true。 
                        false,//    如果要将 PDF 使用范围限制为按照 ISO 19005-1 进行标准化的 PDF 子集，则为 true；否则为 false。如果将此参数设置为 true，则生成的文件会是更加可靠的独立文件，但这些文件可能会更大，或者由于格式限制而显示更多的视觉瑕疵。默认值为 false。
                        oMissing);
                    if (doc != null)
                    {
                        //关闭WORD文件
                        ((WORD._Document)doc).Close(WORD.WdSaveOptions.wdDoNotSaveChanges, Missing.Value, Missing.Value);
                        doc = null;
                    }
                    aPdfFile.Add(sPdfFile);
                }
                return aPdfFile.ToArray();
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                return null;
            }
            finally
            {
                if (applicationClass != null)
                {
                    //退出WORD程序
                    ((WORD._Application)applicationClass).Quit(Missing.Value, Missing.Value, Missing.Value);
                    applicationClass = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        ////使用 Adobe Acrobat DC的OFFICE COM插件,转为PDF
        //public bool PdfMakerToPdf(string fromWordPath, string toPath)
        //{
        //    bool flag = true;
        //    try
        //    {
        //        ///
        //        ///参数：docfile，源office文件绝对路径及文件名（C:\office\myDoc.doc）；printpath，pdf文件保存路径（D:\myPdf）；printFileName，保
        //        ///存pdf文件的文件名（myNewPdf.pdf）
        //        ///
        //        object missing = System.Type.Missing;
        //        PDFMakerApp app = new PDFMakerApp();
        //        var rtnObj = app.CreatePDF(fromWordPath, toPath, PDFMakerSettings.kConvertAllPages, false, false, true, missing);
        //    }
        //    catch (Exception exception)
        //    {
        //        classLims_NPOI.WriteLog(exception, "");
        //        flag = false;
        //    }
        //    return flag;
        //}

        #endregion




    }
    #endregion

}
