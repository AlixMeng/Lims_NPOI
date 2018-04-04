using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using wsdlLib;
using EXCEL = Microsoft.Office.Interop.Excel;
using WORD = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Drawing;

namespace nsLims_NPOI
{
    /// <summary>
    /// 测试用窗体
    /// </summary>
    public partial class Form1 : Form
    {
        /// <summary>
        /// Form1构造函数
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 前台打印输出弹窗
        /// </summary>
        /// <param name="str"></param>
        public static void alert(object str)
        {
            string s = str.ToString();
            if (!s.Equals("System.Int32[]"))
            {
                MessageBox.Show(s);
            }
            else
            {
                s = "";
                int[] intt = (int[])str;
                for (int i = 0; i < intt.Length; i++)
                {
                    s = s + intt[i].ToString() + ", ";
                }
                MessageBox.Show(s);
            }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            string s;

            //classLims_NPOI cln = new classLims_NPOI();
            //ConvertbyPrinter cbp = new ConvertbyPrinter();
            //classExcelMthd cem = new classExcelMthd();
            //ImgConvert ic = new ImgConvert();
            //DocXAction dxa = new DocXAction();
            //MergePDF mpf = new MergePDF();
            //FileConvertClass fcc = new FileConvertClass();
            RegexMatch rm;

            var objs = RegexMatch.RegexUpAndDown("321.1HV_(30.00);322.1HV_(30.00);322.5HV^(30.00);");
            object[] o0 = { "序号", "检测项目",         "分析项",            "样品", "单位",  "技术要求", "实测值",                  "单项结论" };
            object[] o1 = { "1",    "绝缘最薄处厚度",   "绝缘最薄处厚度",    "1#",   "mm",     "≥0.62",  "P^(u1)S-V_(d1)S^(u2)kk",  "P" };
            object[] o2 = { "2",    "导体电阻（20℃）", "导体电阻（20℃）",  "1#",   "Ω/ km", "≤7.41",  "X^(u1)S-C_(d1)S^(u2)mm",  "P" };

            object[] o = { o0, o1, o2 };
            object[] colListC = { "序号", "检测项目", "分析项", "单项结论" };
            object[] sc1 = { ")", "）" };
            object[] sc2 = { "(", "（" };
            object[] specialChars = { sc1, sc2 };
            object[] unpH = { };
            object[] mergeMark = { "0", "0" };

            //object missing = Missing.Value;
            //EXCEL.ApplicationClass excel = null;
            //EXCEL.Workbook wb = null;
            //EXCEL.Workbooks workBooks = null;
            //try
            //{
            //    excel = new EXCEL.ApplicationClass();
            //    excel.DisplayAlerts = false;
            //    workBooks = excel.Workbooks;
            //    wb = workBooks.Open("D:\\默认附页.xls", missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing);

            //    cem.reportFy(wb, 1, o,
            //        colListC, specialChars, unpH,
            //        mergeMark, false,"",false,null);

            //    wb.Save();
            //}
            //catch (Exception ex)
            //{
            //    classLims_NPOI.WriteLog(ex, "");
            //}
            //finally
            //{
            //    if (wb != null)
            //    {
            //        //wb.Close(false, missing, false);
            //        wb.Close(false, missing, missing);
            //        int i = Marshal.ReleaseComObject(wb);
            //        wb = null;
            //    }
            //    if (workBooks != null)
            //    {
            //        workBooks.Close();
            //        int i = Marshal.ReleaseComObject(workBooks);
            //        workBooks = null;
            //    }
            //    if (excel != null)
            //    {
            //        excel.Quit();
            //        int i = Marshal.ReleaseComObject(excel);
            //        excel = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();

            //}

            s = "1234";
            Console.WriteLine(s);

            #region 作废的测试代码

            //string cellValue = "P^(u1)S-V_(d1)S^(u2)kk";//需要用正则表达式匹配的字符串

            //object[] o0 = { "序号", "检测项目", "分析项", "样品", "技术要求",
            //        "单位", "单项结论", "实测值" };
            //object[] o11 = { "1", "抗摆锤冲击能", "抗摆锤冲击能1", "1#", "总量≦50",
            //        "J", "符合", "12.3",};
            //object[] o12 = { "1", "抗摆锤冲击能", "抗摆锤冲击能2", "1#", "总量≦50",
            //        "J", "符合", "14.3",};
            //object[] o13 = { "1", "抗摆锤冲击能", "抗摆锤冲击能3", "1#", "总量≦50",
            //        "J", "符合", "22.3",};
            //object[] o2 = { "2", "耐跌落性（袋）", "耐跌落性（袋）", "1#", "无渗漏，无破裂",
            //        "----", "合格", "c" };
            //object[] o3 = { "3", "甲苯二胺（4%乙酸）", "甲苯二胺（4%乙酸）", "1#", "≤0.004",
            //        "mg/L", "合格", "未检出" };
            //object[] o = { o0, o11,o12,o13, o2, o3 };
            //object[] colListC = { "检测项目", "单位", "单项结论" };
            //object[] sc1 = { "^p", "\n" };
            //object[] sc2 = { "≦", "≤" };
            //object[] sc = { sc1, sc2 };
            //object[] unpH = { };
            //object[] mergeMark = { "1", "0" };
            ////cem.reportFy("D:\\18583溶剂型胶粘剂附页.xls", 1, o, 
            ////    colListC, sc, unpH,
            ////    mergeMark, false);

            //object missing = System.Reflection.Missing.Value;
            //EXCEL.ApplicationClass excel = null;
            //EXCEL.Workbook wb = null;
            //EXCEL.Workbooks workBooks = null;
            //try
            //{
            //    excel = new EXCEL.ApplicationClass();
            //    excel.DisplayAlerts = false;
            //    workBooks = excel.Workbooks;
            //    wb = workBooks.Open("D:\\172边框消失.xls", missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing);

            //    EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Sheets[2];                

            //    string cellValue = classExcelMthd.getMergerCellValue(sheet, 8, 7);
            //    //int upStart = cellValue.IndexOf("^u(");//上标开始标记
            //    //int downStart = cellValue.IndexOf("^d(");//下标开始标记
            //    //int end = cellValue.IndexOf(")");
            //    System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\(([^)]*)\)");
            //    System.Text.RegularExpressions.Match m = reg.Match(cellValue);
            //    if (m.Success)
            //    {
            //        alert(m.Result("$1"));
            //    }
            //    //EXCEL.Range cell = (EXCEL.Range)sheet.Cells[8, 7];
            //    //cell.Characters[3, 1].Font.Superscript = true;//设置上标
            //    //cell.Characters[4, 1].Font.Subscript = true;//设置下标

            //    //wb.Save();
            //}
            //catch (Exception ex)
            //{
            //    classLims_NPOI.WriteLog(ex, "");
            //}
            //finally
            //{
            //    if (wb != null)
            //    {
            //        //wb.Close(false, missing, false);
            //        wb.Close(false, missing, missing);
            //        int i = Marshal.ReleaseComObject(wb);
            //        wb = null;
            //    }
            //    if (workBooks != null)
            //    {
            //        workBooks.Close();
            //        int i = Marshal.ReleaseComObject(workBooks);
            //        workBooks = null;
            //    }
            //    if (excel != null)
            //    {
            //        excel.Quit();
            //        int i = Marshal.ReleaseComObject(excel);
            //        excel = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();

            //}

            //WORD.ApplicationClass applicationClass = null;
            //WORD.Document doc = null;
            //object oFalse = false;
            //object oTrue = true;
            //var oMissing = Type.Missing;
            //try
            //{

            //    applicationClass = new WORD.ApplicationClass();
            //    applicationClass.GetType();
            //    doc = applicationClass.Documents.Open(
            //       "D:\\旧公式.docx",
            //       ref oFalse, //如果该属性为 True，则当文件不是 Microsoft Word 格式时，将显示“转换文件”对话框。
            //       ref oTrue, //如果该属性值为 True，则以只读方式打开文档。
            //       ref oFalse,//如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中
            //       ref oMissing,
            //       ref oMissing,
            //       ref oFalse, //控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。
            //       ref oMissing, ref oMissing,
            //       ref oMissing, ref oMissing, ref oMissing,
            //       ref oMissing, ref oMissing,
            //       ref oMissing, ref oMissing);

            //    var b1 = doc.ActiveWindow.View.ShowDrawings;
            //    var b2 = doc.ActiveWindow.View.DisplayBackgrounds;
            //    flag = true;
            //}
            //catch (Exception exception)
            //{
            //    classLims_NPOI.WriteLog(exception, "");
            //    flag = false;
            //}
            //finally
            //{
            //    if (doc != null)
            //    {
            //        //关闭WORD文件
            //        ((WORD._Document)doc).Close(WORD.WdSaveOptions.wdDoNotSaveChanges, Missing.Value, Missing.Value);

            //        doc = null;
            //    }
            //    if (applicationClass != null)
            //    {
            //        //退出WORD程序
            //        ((WORD._Application)applicationClass).Quit(Missing.Value, Missing.Value, Missing.Value);
            //        applicationClass = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //}

            ////classExcelMthd.ReplaceAll("D:\\默认首页.xls", 1, aSYData);
            //cem.reportStaticExcel("D:\\省级监督抽查封面.xls", 1, aSYData, false, "");

            //object[] aSYData =
            //    { new object[] { "&[样品名称]", "防火材料墙板" },
            //    new object[] { "&[受检单位]", "成都成塑阳光建材有限责任公司" },
            //    new object[] { "&[生产单位]", "成都川立装饰材料有限公司" },
            //    new object[] { "&[委托单位]", "----" },
            //    new object[] { "&[检测类型]", "国家监督专项抽查" },
            //    new object[] { "&[检验站]", "" },
            //    new object[] { "&[检验单位]", "成都市产品质量监督检验院" },
            //    new object[] { "&[任务编号]", "ASHA117Z00002" },
            //    new object[] { "&[标称生产单位]", "标称生产单位" }};

            //object[] imageArray = new object[]
            //{ "D:\\公司 CMA川+CAL川+CNAS.jpg", "D:\\成都质检_带logo.png", "D:\\建设工程专用章.jpg","D:\\reportDownload.png" };
            //object[] aPoint = new object[] {
            //    new object[] { 15, 5, -1, -1 },//资质章
            //    new object[] { 415, 175, 60, 60},//报告下载二维码
            //    new object[] { 15, 80, -1, -1},//建筑方章
            //    new object[] { 15, 175, 60, 60}//建筑二维码
            //};
            //object[] aModelFile = { "D:\\默认封面.xls", "D:\\默认首页.xls", "D:\\3实测值红黄蓝.xls" };
            //object[] aFMData = { new object[] { "&[任务编号]", "ASHA117Z00005" } };
            //object[] aSYData = { new object[] { "&[任务编号]", "ASHA117Z00005" }, new object[] { "&[样品名称]", "FUNCK QIANG" } };
            //object[] aFYData = { new object[] { "&[任务编号]", "ASHA117Z00005" } };
            //object[] o0 = { "序号", "检测项目", "分析项", "样品", "技术要求",
            //        "单位", "单项结论", "实测值1", "实测值2", "实测值3" };
            //object[] o1 = { "1", "抗摆锤冲击能", "抗摆锤冲击能", "1#", "0.8",
            //        "J", "符合", "a", "a", "a" };
            //object[] o2 = { "2", "耐跌落性（袋）", "耐跌落性（袋）", "1#", "无渗漏，无破裂",
            //        "----", "合格", "c", "d", "e" };
            //object[] o3 = { "3", "甲苯二胺（4%乙酸）", "甲苯二胺（4%乙酸）", "1#", "≤0.004",
            //        "mg/L", "合格", "未检出", "未检出", "未检出" };
            //object[] o = { o0, o1, o2, o3 };
            //object[] colListC = { "检测项目", "单位", "单项结论" };
            //object[] sc1 = { ";", "；" };
            //object[] sc2 = { "≦", "≤" };
            //object[] sc = { sc1, sc2 };
            //object[] unpH = { };
            //object[] mergeMark = { "1", "0" };
            //cem.createOneThreeExcelAndMerge(aModelFile,
            //    aFMData, imageArray, aPoint,
            //    aSYData,
            //    aFYData,
            //    o, colListC, sc, unpH, mergeMark,
            //    "D:\\", false);

            //cem.reportCreate_Part("D:\\tt.xls",
            //    new object[] { "D:\\默认封面.xls", "D:\\默认首页.xls", null }, true, false,
            //    new object[] { new object[] { "&[任务编号]", "ASHA117Z00006卡卡卡卡看" } }, new object[] { new object[] { "&[任务编号]", "ASHA117Z00006" } },
            //    null, null, null, null, null, null, false);

            //object[] aModelFile = { "D:\\默认封面.xls", "D:\\默认首页.xls", "D:\\3实测值红黄蓝.xls" };
            //object[] aFMData = { new object[] { "&[任务编号]", "ASHA117Z00005" } };
            //object[] aSYData = { new object[] { "&[任务编号]", "ASHA117Z00005" }, new object[] { "&[样品名称]", "FUNCK QIANG" } };
            //object[] aFYData = { new object[] { "&[任务编号]", "ASHA117Z00005" } };
            //object[] o0 = { "序号", "检测项目", "分析项", "样品", "技术要求",
            //        "单位", "单项结论", "实测值1", "实测值2", "实测值3" };
            //object[] o1 = { "1", "抗摆锤冲击能", "抗摆锤冲击能", "1#", "0.8",
            //        "J", "符合", "a", "a", "a" };
            //object[] o2 = { "2", "耐跌落性（袋）", "耐跌落性（袋）", "1#", "无渗漏，无破裂",
            //        "----", "合格", "c", "d", "e" };
            //object[] o3 = { "3", "甲苯二胺（4%乙酸）", "甲苯二胺（4%乙酸）", "1#", "≤0.004",
            //        "mg/L", "合格", "未检出", "未检出", "未检出" };
            //object[] o = { o0, o1, o2, o3 };
            //object[] colListC = { "检测项目", "单位", "单项结论" };
            //object[] sc1 = { ";", "；" };
            //object[] sc2 = { "≦", "≤" };
            //object[] sc = { sc1, sc2 };
            //object[] unpH = { };
            //object[] mergeMark = { "1", "0" };
            //cem.createOneThreeExcelAndMerge(aModelFile, aFMData, aSYData, aFYData,
            //    o, colListC, sc, unpH, mergeMark,
            //    "D:\\", false);

            //cln.reportCoordinateExcel("D:\\省市监督_一栏首页.xls", 0, "D:\\省市监督_一栏首页.xls",
            //    new object[] { new object[] { "H26", "2017-09-26" } },
            //    new object[] { new object[] { "G8", "PREORDERS.TRADEMARK" },
            //                    new object[] { "G9", "PREORDERS.SPECMODEL" }});

            //int[] inttt = new int[] { 1, 2, 3, 4, 8 };
            //object missing = System.Reflection.Missing.Value;
            //EXCEL.ApplicationClass excel = null;
            //EXCEL.Workbook wb = null;
            //EXCEL.Workbooks workBooks = null;
            //try
            //{
            //    excel = new EXCEL.ApplicationClass();
            //    excel.DisplayAlerts = false;
            //    workBooks = excel.Workbooks;
            //    wb = workBooks.Open("D:\\省市监督_一栏首页.xls", missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing);
            //    //实例化Sheet后,释放Excel进程就会失败
            //    //对于sheet的操作必须放在新的方法中,接口层级为Workbook
            //    //excel.ActiveWindow.View = EXCEL.XlWindowView.xlPageBreakPreview;
            //    cem.protectWorkSheet(wb, 1, "111",
            //        true, true, true,
            //        false, false, false,
            //        true,/*允许设置行格式,拉伸行*/
            //false, false,
            //        false, false, false,
            //        false, false, false);
            //    //再还原为普通视图
            //    //excel.ActiveWindow.View = EXCEL.XlWindowView.xlNormalView;
            //    wb.Save();
            //}
            //catch (Exception ex)
            //{
            //    classLims_NPOI.WriteLog(ex, "");
            //}
            //finally
            //{
            //    if (wb != null)
            //    {
            //        //wb.Close(false, missing, false);
            //        wb.Close(false, missing, missing);
            //        int i = Marshal.ReleaseComObject(wb);
            //        wb = null;
            //    }
            //    if (workBooks != null)
            //    {
            //        workBooks.Close();
            //        int i = Marshal.ReleaseComObject(workBooks);
            //        workBooks = null;
            //    }
            //    if (excel != null)
            //    {
            //        excel.Quit();
            //        int i = Marshal.ReleaseComObject(excel);
            //        excel = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();

            //}

            //object[] o0 = { "序号", "检测项目", "分析项", "样品", "技术要求",
            //    "单位", "单项结论", "实测值1", "实测值2", "实测值3" };
            //object[] o1 = { "1", "抗摆锤冲击能", "抗摆锤冲击能", "1#", "0.8",
            //    "J", "符合", "a", "a", "a" };
            //object[] o2 = { "2", "耐跌落性（袋）", "耐跌落性（袋）", "1#", "无渗漏，无破裂",
            //    "----", "合格", "c", "d", "e" };
            //object[] o3 = { "3", "甲苯二胺（4%乙酸）", "甲苯二胺（4%乙酸）", "1#", "≤0.004",
            //    "mg/L", "合格", "未检出", "未检出", "未检出" };
            //object[] o = { o0, o1, o2, o3 };
            //object[] colListC = { "检测项目", "单位", "单项结论" };
            //object[] sc1 = { ";", "；" };
            //object[] sc2 = { "≦", "≤" };
            //object[] sc = { sc1, sc2 };
            //object[] unpH = { };
            //object[] mergeMark = { "1", "0" };
            //cln.reportOneDimDExcel("D:\\3实测值红黄蓝.xls", 0, "D:\\3实测值红黄蓝.xls", o, colListC, 9.25, sc, unpH, mergeMark);

            //object[] o1 = { ")", "）" };
            //object[] o2 = { "(", "（" };
            //object[] o3 = { "%", "％" };
            //object[] o = { o1, o2, o3 };
            //cem.reportOneDimDExcelFormat("D:\\TEST.xls", 1, new int[] { 1, 2, 3, 8 }, 5, 9.25, 1, 8, o);

            //mpf.addImageToPdf_Left("D:\\打印版.pdf", "D:\\打印版_new.pdf", "D:\\公司 CMAF川+CNAS.jpg", 93.5, 700.0, 100);
            //mpf.addImageToPdf("D:\\打印版_new.pdf", "D:\\打印版.pdf", "D:\\建设工程专用章.jpg", 120.0, 745, 146.2, "");
            //mpf.addImageToPdf("D:\\打印版.pdf", "D:\\打印版_new.pdf", "D:\\建设工程专用章.jpg", 100.0, 625, 146.2, "");


            //建筑方章
            //mpf.addImageToPdf_Left("D:\\默认封面.pdf", "D:\\默认封面_new.pdf", "D:\\公司 CMAF川+CNAS.jpg", 93.5, 670.0, 100);
            //mpf.addImageToPdf("D:\\默认封面_new.pdf", "D:\\默认封面.pdf", "D:\\建设工程专用章.jpg", 120.0, 745, 146.2, "");

            //object missing = System.Reflection.Missing.Value;
            //string strTargetFile = "";
            //EXCEL.ApplicationClass excel = null;
            //EXCEL.Workbook wb = null;
            //EXCEL.Workbooks workBooks = null;
            //try
            //{
            //    excel = new EXCEL.ApplicationClass();
            //    excel.DisplayAlerts = false;
            //    workBooks = excel.Workbooks;
            //    wb = workBooks.Open("D:\\测试行高.xlsx", missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing, missing, missing, missing,
            //        missing, missing);
            //    //实例化Sheet后,释放Excel进程就会失败
            //    //对于sheet的操作必须放在新的方法中,接口层级为Workbook
            //    double headH = 0;
            //    EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[1];
            //    EXCEL.Range range = (EXCEL.Range)sheet.Rows[1];
            //    headH = (double)range.Height;
            //    alert(headH);
            //    range = (EXCEL.Range)sheet.Rows[2];
            //    headH = (double)range.Height;
            //    alert(headH);
            //}
            //catch (Exception ex)
            //{
            //    classLims_NPOI.WriteLog(ex, "");
            //}
            //finally
            //{
            //    if (wb != null)
            //    {
            //        //wb.Close(false, missing, false);
            //        wb.Close(false, missing, missing);
            //        int i = Marshal.ReleaseComObject(wb);
            //        wb = null;
            //    }
            //    if (workBooks != null)
            //    {
            //        workBooks.Close();
            //        int i = Marshal.ReleaseComObject(workBooks);
            //        workBooks = null;
            //    }
            //    if (excel != null)
            //    {
            //        excel.Quit();
            //        int i = Marshal.ReleaseComObject(excel);
            //        excel = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();

            //}

            //object[] o1 = { "&{主检}", "D:\\ding.du.png" };
            //object[] o2 = { "&{二审}", "D:\\internet.jpg" };
            //object[] o3 = { "&{终审}", "D:\\建设工程专用章.jpg" };
            //object[] o = { o1, o2, o3 };
            //cem.addImagesToExcel_byOffice("D:\\默认首页.xls", 0, o, 63, 24);

            //mpf.addPagenoToOnePage("D:\\sSyFy.pdf", "D:\\waterpdf.pdf", 550, 727, 1);
            //dxa.InsertPicture("D:\\Jl5020102.docx", "&[评定人]", "D:\\dd.jpg", "RIGHT", 40.0, 100.0);
            //mpf.replaceOnePage("D:\\TEST_final.pdf", 3, "D:\\默认首页.pdf", "D:\\replacedPdf.pdf");
            //cem.addImage2Excel_byOffice("D:\\默认首页.xls", 0, "D:\\ding.du.png", "&{主检}", 63, 21);
            //mpf.GetOnepage("D:\\TEST_final.pdf", 3, "D:\\TEST_0000.pdf");
            //ic.addImage2Pdf_path("D:\\TEST_0000.pdf", "D:\\TEST_111_signed.pdf", "&{主检}", "D:\\ding.du.png", (float)63.0, (float)21.0);
            //ic.addIssueDateToPdf("D:\\TEST_final.pdf", "D:\\TEST_0000.pdf", "签发日期：", "2017-06-18");
            //mpf.InsertPageToPdf("D:\\TEST2.pdf", 2, "D:\\reportNull.pdf", 1, "D:\\TEST3.pdf");
            //mpf.MergeAttachments("D:\\TEST_A - 副本.pdf,D:\\TEST_A.pdf", "D:\\TEST_merge22.pdf");


            //cln.dealMergedAreaInPages_new("D:\\TEST.xls", 0);
            //string filePath = "D:\\TEST.xls";
            //IWorkbook wb = cln.loadExcelWorkbookI(filePath);
            //ISheet sheet = wb.GetSheetAt(0);
            //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(36, 36, 5, 6));
            //cln.saveExcelWithoutAsk(filePath, wb);

            //object[] oh = { "序号", "技术要求", "检测项目", "分析项", "实测值", "单位", "单项结论", "样品" };
            //object[] o1 = { 1, ">20", "OO颜色颜色颜色颜色颜色颜色颜色颜色颜色颜色KK", "OO颜色颜色颜色颜色颜色颜色颜色颜色颜色颜色KK", 2.20, "cm", "合格", "001" };
            //object[] o2 = { 2, ">10", "外观色差", "oo对比度对比度对比度对比度对比度对比度对kk", 2.20, "cm", "合格", "001" };
            //object[] o3 = { 3, ">30", "CCCo", "10001001", 2.20, "cm*pasc/mg*L", "合格", "002" };
            //object[] o4 = { 3, ">20", "CCCo", "2233", 2.20, "cm", "合格", "002" };
            //object[] o = { oh, o1, o2, o3, o4, };
            //object[] colListC = { "检测项目", "单位", "单项结论" };

            //cln.reportOneDimDExcel("D:\\默认附页.xls", 0, "D:\\默认附页_new.xls", o, colListC, 9.75);
            //cln.stretchLastRowHeight("D:\\默认附页_new.xls", 0);

            //FileConvertClass fcc = new FileConvertClass();
            //MergePDF mpf = new MergePDF();
            //DocXAction dxa = new DocXAction();
            //classExcelMthd cem = new classExcelMthd();
            //PdfSignGJ psg = new PdfSignGJ();
            //string strMsg = psg.KeySign("检验报告专用章", "D:\\biceng\\pdf\\170117191241_AXFA1W170117004.pdf", "D:\\biceng", "GZ", "123", "D:\\biceng\\pdf", "test_signed");
            //alert(strMsg);
            //string[] flags = { "&[样品编号]", "&[备注]" };
            //cln.protectExcel("D:\\默认首页.xls", 0, flags, "123");


            //cem.protectWorkBook("D:\\TEST.xls", "123");
            //cln.protectSheets("D:\\TEST.xls", "123");
            //cln.stretchLastRowHeight("D:\\TEST.xls", 0);
            //ISheet sheet = cln.loadExcelSheetI("D:\\TEST.xls", 0);
            //List<int> list = cln.getNewPageFirstRow(sheet);            
            //bool isnew = cln.IsNewPageRow(sheet, 113);

            //cln.dealMergedAreaInPages("D:\\TEST.xls", 0);
            //cln.stretchLastRowHeight("D:\\TEST.xls", 0);

            //IWorkbook wb = cln.loadExcelWorkbookI("D:\\默认封面.xls");
            //classLims_NPOI.LockSheet(wb, 0, "lims@123");
            //wb.GetSheetAt(0).GetRow(22).GetCell(0).SetCellValue("成都产品质量检验研究院有限责任公司");
            //wb.GetSheetAt(0).GetRow(23).GetCell(0).SetCellValue("(四川省产品质量监督检验检测院/");
            //wb.GetSheetAt(0).ForceFormulaRecalculation = true;//计算Excel公式
            //cln.saveExcelWithoutAsk("D:\\默认封面_add.xls", wb);            
            //fcc.addImage2Excel_byOffice("D:\\默认封面.xls", 0, "D:\\label2.png", "&[二维码]", 47.5, 47.5);
            //FileConvertClass.excelRefresh("D:\\默认封面.xls");

            //object[] obj1 = { "20161117080005", "20161117080005", "20161117080005" };
            //object[] obj2 = { "001", "001", "001" };
            //object[] obj3 = { "10007508", "10007509", "10007510" };
            //object[] obj4 = { "img1:化石研究", "img2:布偶猫2号", "img3:布偶猫3号" };
            //object[] obj5 = { "D:\\1化石研究.jpg", "D:\\2cat.jpg", "D:\\2cat.jpg" };

            //IWorkbook wb = cln.loadExcelWorkbookI("D:\\默认附页2.xls");
            //wb = cln.reportImagesExcel(wb, 1, obj1, obj2, obj3, obj4, obj5, "— — — — 以下空白 — — — —");
            //cln.saveExcelWithoutAsk("D:\\默认附页2 - 副本.xls", wb);

            //object[] obj1 = { "&[任务编号]", "20161117080005" };
            //object[] obj2 = { "&[报告页数]", "4" };
            //object[] obj3 = { "&[当前页数]", "4" };
            //object[] obj4 = { "&[子样1]", "001" };
            //object[] obj5 = { "&[检测项1]", "10007508" };
            //object[] obj6 = { "&[图片说明1]", "img1:布偶猫1号" };
            //object[] obj = { obj1, obj2, obj3, obj4, obj5, obj6 };

            ////cln.addImgTo2ImgWorkbook("D:\\2图片模板 - 副本.xls", 0, "D:\\2cat.jpg",
            ////    "", "&[图片1]", "&[图片2]", "— — — — 以下空白 — — — —", obj);
            //cln.addImgTo1ImgWorkbook("D:\\1图片模板 - 副本.xls", 0, "D:\\1化石研究.jpg",
            //    "&[图片1]", obj);            

            //cbp.ConvertExcelWorkSheetPDF_index("D:\\常用密码管理.xlsx", "D:\\常用密码管理.pdf", 1);

            //cbp.ConvertExcelWorkSheetPDF("D:\\1111.xlsx", "D:\\1111.pdf", "Sheet1");


            //ic.addImage2Pdf_path("D:\\excel.pdf", "D:\\word_01.pdf", "&{主检}", "D:\\ZZY.png", 90, 30);
            //string excelPath = "D:\\1234.xls";
            //string toPath = "D:\\默认附页_add.xls";

            //HSSFWorkbook wb = cln.loadExcelWorkbook(excelPath);
            //HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);
            //cln.stretchLastRowHeight(excelPath, 0, 1, 2, 0, 7);
            //cbp.ConvertExcelWorkSheetPDF(excelPath, toPath, "Sheet1");
            //cln.saveExcelWithoutAsk(toPath, wb);




            //string excelPath = "D:\\1实测值.xls";
            //string toPath = "D:\\1实测值_add.xls";

            //HSSFWorkbook wb = cln.loadExcelWorkbook(excelPath);
            //HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);
            //cln.stretchLastRowHeight(sheet, 1, 2, 0, 7);
            //cln.saveExcelWithoutAsk(toPath, wb);首页

            //object[] obj1 = { "&[任务编号]", "报告书编号：任务编号" };
            //object[] obj2 = { "&[检验单位]", "检验单位名称：检验单位" };
            //object[] obj = { obj1, obj2 };

            //object[] oh = { "序号", "技术要求", "检测项目", "分析项", "实测值", "单位", "单项结论", "小样" };
            //object[] o1 = { 1, ">20", "外观", "颜色", 2.20, "cm", "合格", "001" };
            //object[] o2 = { 2, ">10", "外观", "对比度", 2.20, "cm", "合格", "001" };
            //object[] o3 = { 3, ">30", "CCCo", "10001001", 2.20, "cm*pasc/mg*L", "合格", "002" };
            //object[] o4 = { 4, ">20", "CCCo", "2233", 2.20, "cm", "合格", "002" };
            //object[] o = { oh, o1, o2, o3, o4, };

            //object[] colListC = { "检测项目", "单项结论" };


            #endregion

            button1.Enabled = true;
        }
        
    }
}
