﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using wsdlLib;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace nsLims_NPOI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static void alert(object str)
        {
            string s = str.ToString();
            MessageBox.Show(s);
        }
       

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;

            //classLims_NPOI cln = new classLims_NPOI();
            //ConvertbyPrinter cbp = new ConvertbyPrinter();
            classExcelMthd cem = new classExcelMthd();
            //ImgConvert ic = new ImgConvert();
            //DocXAction dxa = new DocXAction();
            //MergePDF mpf = new MergePDF();

            object missing = Type.Missing;
            EXCEL.ApplicationClass excel = null;
            EXCEL.Workbook wb = null;
            EXCEL.Workbooks workBooks = null;
            try
            {
                excel = new EXCEL.ApplicationClass();
                workBooks = excel.Workbooks;
                wb = workBooks.Open("D:\\读取行高.xlsx", missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing);
                //cem.dealMergedAreaInPages_new(wb, 1, 39, 8);
                EXCEL.Worksheet sheet = (EXCEL.Worksheet)wb.Worksheets[1];
                for (int i = 1; i <= 2; i++)
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
                }
                wb.SaveAs("D:\\读取行高_new.xls", EXCEL.XlFileFormat.xlExcel8, null, null, false, false, EXCEL.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
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


            string s = "";
            #region

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
