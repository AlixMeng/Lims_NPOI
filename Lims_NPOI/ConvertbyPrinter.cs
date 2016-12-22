using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;


namespace nsLims_NPOI
{
    /// <summary>
    /// 使用打印功能转换文档
    /// </summary>
    class ConvertbyPrinter
    {

        #region 系统变量

        private string _sourcePath = "";

        private string _targetPath = "";

        private string _worksheetName = "";

        private int _worksheetIndex = 0;

        public int worksheetIndex
        {
            get
            {
                return this._worksheetIndex;
            }
            set
            {
                this._worksheetIndex = value;
            }
        }

        public string sourcePath
        {
            get
            {
                return this._sourcePath;
            }
            set
            {
                this._sourcePath = value;
            }
        }

        public string targetPath
        {
            get
            {
                return this._targetPath;
            }
            set
            {
                this._targetPath = value;
            }
        }

        public string worksheetName
        {
            get
            {
                return this._worksheetName;
            }
            set
            {
                this._worksheetName = value;
            }
        }

        #endregion

        public ConvertbyPrinter()
        {
        }
       
        /// <summary>
        /// 杀死进程,已用于杀死pdf打印机进程,请谨慎使用
        /// </summary>
        /// <param name="processName">进程名,如"pdfSaver5"</param>
        public void KillProcess(string processName) //调用方法，传参
        {
            try
            {
                System.Diagnostics.Process[] thisproc = System.Diagnostics.Process.GetProcessesByName(processName);
                //thisproc.lendth:名字为进程总数
                if (thisproc.Length > 0)
                {
                    for (int i = 0; i < thisproc.Length; i++)
                    {
                        if (!thisproc[i].CloseMainWindow()) //尝试关闭进程 释放资源
                        {
                            thisproc[i].Kill(); //强制关闭
                            return;
                        }
                        
                    }
                }
            }
            catch (Exception ex) //出现异常，表明 kill 进程失败
            {
                classLims_NPOI.WriteLog(ex, "");
            }
        }


        #region 打印excel

        
        /// <summary>
        /// 按名称查找sheet,使用默认pdf打印机转换为pdf
        /// </summary>
        public bool ConvertExcelWorkSheetPDF()
        {
            try
            {
                bool flag = true;
                if (this._sourcePath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的源文件路径不能为空。");
                }
                //if (this._targetPath.Length == 0)
                //{
                //    flag = false;
                //    throw new Exception("转换后文件路径不能为空。");
                //}
                if (this._worksheetName.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的WorkSheet不能为空。");
                }
                Microsoft.Office.Interop.Excel.ApplicationClass applicationClass = new Microsoft.Office.Interop.Excel.ApplicationClass();
                applicationClass.GetType();
                Workbooks workbooks = applicationClass.Workbooks;//.get_Workbooks();
                Type type = workbooks.GetType();
                object obj = this._sourcePath;
                object[] objArray = new object[] { obj, true, true };
                Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)type.InvokeMember("Open", BindingFlags.InvokeMethod,
                    null, workbooks, objArray);
                //workbooks.Open(this._sourcePath);
                //Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Item[0];
                workbook.GetType();
                //object obj1 = "c:\\temp.ps";
                //obj1 = this._targetPath;
                object obj2 = "-4142";
                Microsoft.Office.Interop.Excel.Worksheet item = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Item[this._worksheetName];
                if (item == null)
                {
                    flag = false;
                    throw new Exception("需要转换的WorkSheet名称不存在。");
                }
                //item.get_Cells().get_Interior().set_ColorIndex(obj2);
                item.Cells.Interior.ColorIndex = obj2;//.get_Interior().set_ColorIndex(obj2);
                object value = Missing.Value;
                //item.PrintOut(value, value, value, value, value, true, value, obj1);
                //Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                //item.ExportAsFixedFormat(targetType, obj1, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,true,false, value, value, value, value);

                //目标路径仅在打印失败时写入,成功时都默认在打印机路径下
                //故不使用目标路径,直接使用打印机默认路径
                item.PrintOut(value, value, value, value, value, false, value, value);


                if (item != null)
                {

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    //Marshal.FinalReleaseComObject(sheet);
                    item = null;
                }
                if (workbook != null)
                {
                    workbook.Close(false, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    //Marshal.FinalReleaseComObject(workBook);
                    workbook = null;
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(applicationClass);
                    applicationClass = null;

                }



                GC.Collect();
                GC.WaitForPendingFinalizers();
                return flag;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return false;
            }
        }

        
        /// <summary>
        /// 按名称查找sheet,使用默认pdf打印机转换为pdf
        /// </summary>
        /// <param name="fromExcelPath">源excel路径</param>
        /// <param name="worksheetName">sheet名</param>
        /// <returns>成功或失败</returns>
        public bool ConvertExcelWorkSheetPDF(string fromExcelPath, string worksheetName)
        {
            bool flag = false;
            this._sourcePath = fromExcelPath;
            this._worksheetName = worksheetName;
            try
            {
                try
                {
                    flag = this.ConvertExcelWorkSheetPDF();
                    //flag = this.ConvertExcelWorkSheetPDF_pdfAction();
                }
                catch (Exception exception)
                {
                    classLims_NPOI.WriteLog(exception, "");
                    throw exception;
                }
            }
            finally
            {
            }
            return flag;
        }

        
        /// <summary>
        /// 按索引查找sheet,使用默认pdf打印机转换为pdf
        /// </summary>
        /// <param name="fromExcelPath">源excel路径</param>
        /// <param name="worksheetIndex">sheet索引,注意从1开始</param>
        /// <returns>成功或失败</returns>
        public bool ConvertExcelWorkSheetPDF_index(string fromExcelPath, int worksheetIndex)
        {
            bool flag = true;

            try
            {
                if (fromExcelPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的源文件路径不能为空。");
                }
                Microsoft.Office.Interop.Excel.ApplicationClass applicationClass = new Microsoft.Office.Interop.Excel.ApplicationClass();
                applicationClass.GetType();
                Workbooks workbooks = applicationClass.Workbooks;//.get_Workbooks();
                Type type = workbooks.GetType();
                object obj = fromExcelPath;
                object[] objArray = new object[] { obj, true, true };
                Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)type.InvokeMember("Open", BindingFlags.InvokeMethod,
                    null, workbooks, objArray);
                //workbooks.Open(this._sourcePath);
                //Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Item[0];
                workbook.GetType();
                //object obj1 = "c:\\temp.ps";
                //obj1 = toPdfPath;
                object obj2 = "-4142";
                //Microsoft.Office.Interop.Excel.Worksheet item = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Item[worksheetIndex];
                Microsoft.Office.Interop.Excel.Worksheet item = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(worksheetIndex);
                if (item == null)
                {
                    flag = false;
                    throw new Exception("需要转换的WorkSheet名称不存在。");
                }
                //item.get_Cells().get_Interior().set_ColorIndex(obj2);
                item.Cells.Interior.ColorIndex = obj2;//.get_Interior().set_ColorIndex(obj2);
                object value = Missing.Value;
                //item.PrintOut(value, value, value, value, value, true, value, obj1);
                //Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                //item.ExportAsFixedFormat(targetType, obj1, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,true,false, value, value, value, value);

                //目标路径仅在打印失败时写入,成功时都默认在打印机路径下
                //故不使用目标路径,直接使用打印机默认路径
                item.PrintOut(value, value, value, value, value, false, value, value);


                if (item != null)
                {

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    //Marshal.FinalReleaseComObject(sheet);
                    item = null;
                }
                if (workbook != null)
                {
                    workbook.Close(false, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    //Marshal.FinalReleaseComObject(workBook);
                    workbook = null;
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(applicationClass);
                    applicationClass = null;

                }



                GC.Collect();
                GC.WaitForPendingFinalizers();
                return flag;
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                throw exception;
            }
            
            finally
            {
            }
            
        }

        //打印整个workbook
        /// <summary>
        /// 打印整个workbook,使用默认pdf打印机转换为pdf
        /// </summary>
        /// <param name="fromExcelPath">源excel路径</param>
        /// <returns>成功或失败</returns>
        public bool ConvertExcelWorkbookPDF(string fromExcelPath)
        {
            bool flag = true;

            try
            {
                if (fromExcelPath.Length == 0)
                {
                    flag = false;
                    throw new Exception("需要转换的源文件路径不能为空。");
                }
                Microsoft.Office.Interop.Excel.ApplicationClass applicationClass = new Microsoft.Office.Interop.Excel.ApplicationClass();
                applicationClass.GetType();
                Workbooks workbooks = applicationClass.Workbooks;//.get_Workbooks();
                Type type = workbooks.GetType();
                object obj = fromExcelPath;
                object[] objArray = new object[] { obj, true, true };
                Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)type.InvokeMember("Open", BindingFlags.InvokeMethod,
                    null, workbooks, objArray);

                workbook.GetType();
                object value = Missing.Value;

                //目标路径仅在打印失败时写入,成功时都默认在打印机路径下
                //故不使用目标路径,直接使用打印机默认路径
                workbook.PrintOutEx(value, value, value, value, value, false, value, value, value);
                //item.PrintOut(value, value, value, value, value, false, value, value);                 

                if (workbook != null)
                {
                    workbook.Close(false, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    //Marshal.FinalReleaseComObject(workBook);
                    workbook = null;
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(applicationClass);
                    applicationClass = null;

                }



                GC.Collect();
                GC.WaitForPendingFinalizers();
                return flag;
            }
            catch (Exception exception)
            {
                classLims_NPOI.WriteLog(exception, "");
                throw exception;
            }

            finally
            {
            }

        }
        
       
        #endregion

        #region 打印word        

        /// <summary>
        /// 使用Word.Application的打印功能
        /// </summary>
        /// <returns></returns>
        public bool ConvertWord2PDF()
        {
            try
            {
                bool flag = true;
                object wordFile = this._sourcePath;//@"c:\test.doc";
                object oMissing = Missing.Value;

                //自定义object类型的布尔值
                object oTrue = true;
                object oFalse = false;
                object Copies = 1; //打印份数
                object wdPrintFrom = 1;//打印的起始页码
                object wdPrintTo = 1;//打印的结束页码

                object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;

                //定义WORD Application相关
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();

                //WORD程序不可见
                appWord.Visible = false;
                //不弹出警告框
                appWord.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;

                //先保存默认的打印机
                string defaultPrinter = appWord.ActivePrinter;

                //打开要打印的文件
                Microsoft.Office.Interop.Word.Document doc = appWord.Documents.Open(
                    ref wordFile,
                    ref oMissing,
                    ref oTrue,
                    ref oFalse,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //设置指定的打印机
                appWord.ActivePrinter = defaultPrinter; //"\\\\10.10.96.236\\HP Deskjet 6500 Series";

                //打印
                doc.PrintOut(
                    ref oTrue, //Background 此处为true,表示后台打印
                    ref oFalse,
                    ref oMissing, //Range 页面范围
                    ref oMissing, //this._targetPath,//打印路径
                    ref oMissing, //当 Range 设置为 wdPrintFromTo 时的起始页码
                    ref oMissing,//当 Range 设置为 wdPrintFromTo 时的结束页码
                    ref oMissing,
                    ref Copies,  //要打印的份数
                    ref oMissing, ref oMissing,
                    ref oMissing,//是否打印到文件
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);



                //打印完关闭WORD文件
                ((Microsoft.Office.Interop.Word._Document)doc).Close(ref doNotSaveChanges, ref oMissing, ref oMissing);

                //还原原来的默认打印机
                appWord.ActivePrinter = defaultPrinter;

                //退出WORD程序
                ((Microsoft.Office.Interop.Word._Application)appWord).Quit(ref oMissing, ref oMissing, ref oMissing);

                doc = null;
                appWord = null;


                return flag;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return false;
            }
        }
        
        /// <summary>
        /// 使用Word.Application的打印功能
        /// </summary>
        /// <param name="strSourcePath">源路径</param>
        /// <returns></returns>
        public bool ConvertWord2PDF(string strSourcePath)
        {
            bool flag = false;
            this._sourcePath = strSourcePath;
            try
            {
                try
                {
                    flag = this.ConvertWord2PDF();
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
            finally
            {
            }
            return flag;
        }

       

        #endregion




        //通过移动方法,检查打印任务是否结束
        /// <summary>
        /// 通过移动方法,检查打印任务是否结束
        /// </summary>
        /// <param name="fromFile">源文件全路径</param>
        /// <param name="toFile">目标文件路径</param>
        /// <returns>是否移动成功</returns>
        public bool fileMove(string fromFile, string toFile)
        {
            try
            {
                FileInfo fi = new FileInfo(toFile);
                var di = fi.Directory;
                if (!di.Exists)
                    di.Create();
                //一直检查,看是否到了保存这一步
                while (!File.Exists(fromFile))
                {
                    ;
                }
                //通过移动函数,检查打印进程是否保存文件结束
                System.IO.File.Move(fromFile, toFile);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void printPDF()
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            process.StartInfo.UseShellExecute = true;
            process.StartInfo.FileName = this._sourcePath;
            process.StartInfo.Verb = "print";
            process.Start();
        }

        public bool printPDF(string strSourcePath)
        {
            this._sourcePath = strSourcePath;
            try
            {
                try
                {
                    this.printPDF();
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
            finally
            {
            }
            return true;
        }
    }

}