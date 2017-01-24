using System;
using System.IO;
using System.Runtime.InteropServices;
using EXCEL = Microsoft.Office.Interop.Excel;
using WORD = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.VisualBasic.Devices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace nsLims_NPOI
{
    class classExcelMthd
    {

        //向指定sheet添加图片并保存
        private string addImageToSheet(string workbookPath, EXCEL.Workbook wb, int sheetIndex, string rangeName,
            string imagePath, string imgFlag, double PicWidth, double PicHeight)
        {
            object missing = System.Reflection.Missing.Value;
            EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex + 1];//索引从1开始
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

            string newGUID = System.Guid.NewGuid().ToString();
            string strTargetFile = workbookPath.Substring(0, workbookPath.LastIndexOf("\\") + 1)
                + newGUID
                + workbookPath.Substring(workbookPath.LastIndexOf("."));
            wb.SaveAs(strTargetFile, wb.FileFormat, missing, missing, missing, missing, EXCEL.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);
            return strTargetFile;
        }

        /// <summary>
        /// 添加图片到指定excel,图片指定大小,使用office的com组件
        /// </summary>
        /// <param name="workbookPath">源工作簿路径</param>
        /// <param name="sheetIndex">工作表sheet索引</param>
        /// <param name="toPath">excel保存路径</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="imgFlag">图片文件路径</param>
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
                //sheet = (EXCEL.Worksheet)wb.Worksheets[sheetIndex + 1];//索引从1开始

                strTargetFile = addImageToSheet(workbookPath, wb, sheetIndex, rangeName, imagePath, imgFlag, PicWidth, PicHeight);

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
            string strOldFileName = workbookPath.Substring(workbookPath.LastIndexOf("\\") + 1, 
                workbookPath.Length - workbookPath.LastIndexOf("\\") - 1);
            if (File.Exists(strTargetFile))
            {
                File.Delete(workbookPath);
                Computer MyComputer = new Computer();
                MyComputer.FileSystem.RenameFile(strTargetFile, strOldFileName);
                flag = true;
            }
            else
            {
                flag = false;
            }
            return flag;

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
                //string newFile = System.Guid.NewGuid().ToString();
                //string strTargetFile = strSourceFile.Substring(0, strSourceFile.LastIndexOf("\\") + 1)
                //    + newFile.ToString()
                //    + strSourceFile.Substring(strSourceFile.LastIndexOf("."));
                //string targetFile = strTargetFile;
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
                //string strOldFileName = strSourceFile.Substring(strSourceFile.LastIndexOf("\\") + 1, strSourceFile.Length - strSourceFile.LastIndexOf("\\") - 1);
                //File.Delete(strSourceFile);
                //Computer MyComputer = new Computer();
                //MyComputer.FileSystem.RenameFile(strTargetFile, strOldFileName);
            }
            return flag;
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


    }
}
