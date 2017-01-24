using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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

            classLims_NPOI cln = new classLims_NPOI();
            //ConvertbyPrinter cbp = new ConvertbyPrinter();
            //ImgConvert ic = new ImgConvert();
            //FileConvertClass fcc = new FileConvertClass();
            //MergePDF mpf = new MergePDF();
            //DocXAction dxa = new DocXAction();
            //classExcelMthd cem = new classExcelMthd();

            string[] flags = { "&[样品编号]", "&[备注]" };
            cln.protectExcel("D:\\默认首页.xls", 0, flags, "123");



            #region
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
