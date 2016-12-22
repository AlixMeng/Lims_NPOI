using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;

namespace nsLims_NPOI
{
    class ELNtoPDF
    {
        private const int PAGE_WIDTH_OFFSET = 30;

        private const int PAGE_HEIGHT_OFFSET = 10;

        private const int CONTENT_TOP_OFFSET = 50;

        private const int PAGE_HEADER_HEIGHT = 50;

        private string _tester;

        private string _testDateTime;

        private string _corrector;

        private string _correctDateTime;

        private string _auditor;

        private string _auditDateTime;

        public string AuditDateTime
        {
            get
            {
                return this._auditDateTime;
            }
            set
            {
                this._auditDateTime = value;
            }
        }

        public string Auditor
        {
            get
            {
                return this._auditor;
            }
            set
            {
                this._auditor = value;
            }
        }

        public string CorrectDateTime
        {
            get
            {
                return this._correctDateTime;
            }
            set
            {
                this._correctDateTime = value;
            }
        }

        public string Corrector
        {
            get
            {
                return this._corrector;
            }
            set
            {
                this._corrector = value;
            }
        }

        public string TestDateTime
        {
            get
            {
                return this._testDateTime;
            }
            set
            {
                this._testDateTime = value;
            }
        }

        public string Tester
        {
            get
            {
                return this._tester;
            }
            set
            {
                this._tester = value;
            }
        }

        public ELNtoPDF()
        {
        }

        public bool CreatePageFlag(string strPDFPath, string strNewPDFPath)
        {
            bool flag = false;
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(strNewPDFPath, FileMode.Create));
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    PdfReader pdfReader = new PdfReader(strPDFPath);
                    int numberOfPages = pdfReader.NumberOfPages;
                    for (int i = 1; i <= numberOfPages; i++)
                    {
                        document.NewPage();
                        if ((int)pdfReader.GetPageContent(i).Length > 0)
                        {
                            directContent.AddTemplate(instance.GetImportedPage(pdfReader, i), 0f, 0f);
                            if (numberOfPages != 1)
                            {
                                if (i == 1)
                                {
                                    directContent.BeginText();
                                    this.CreatePDFNext(directContent);
                                    directContent.EndText();
                                }
                                else if (i != numberOfPages)
                                {
                                    directContent.BeginText();
                                    this.CreatePDFNext(directContent);
                                    this.CreatePDFPrevious(directContent);
                                    directContent.EndText();
                                }
                                else
                                {
                                    directContent.BeginText();
                                    this.CreatePDFPrevious(directContent);
                                    directContent.EndText();
                                }
                            }
                        }
                    }
                    flag = true;
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
            finally
            {
                if (document.IsOpen())
                {
                    document.Close();
                }
            }
            return flag;
        }

        public bool CreatePDF(string strPDFPath, string strNewPDFPath)
        {
            bool flag = false;
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(strNewPDFPath, FileMode.Create));
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    PdfReader pdfReader = new PdfReader(strPDFPath);
                    int numberOfPages = pdfReader.NumberOfPages;
                    for (int i = 1; i <= numberOfPages; i++)
                    {
                        document.NewPage();
                        if ((int)pdfReader.GetPageContent(i).Length > 0)
                        {
                            directContent.AddTemplate(instance.GetImportedPage(pdfReader, i), 0f, 0f);
                        }
                    }
                    directContent.BeginText();
                    this.CreatePDFFooter(directContent);
                    directContent.EndText();
                    flag = true;
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
            finally
            {
                if (document.IsOpen())
                {
                    document.Close();
                }
            }
            return flag;
        }

        private void CreatePDFBLANKFooter(PdfContentByte pdfContentByte)
        {
            Color color = new Color(255, 255, 255);
            pdfContentByte.SetColorFill(color);
            pdfContentByte.SetColorStroke(color);
            pdfContentByte.SetLineWidth(20f);
            pdfContentByte.MoveTo(30f, 876f);
            pdfContentByte.LineTo(PageSize.A4.Width - 30f, 876f);
            pdfContentByte.MoveTo(0f, 10f);
            pdfContentByte.LineTo(PageSize.A4.Width, 10f);
            pdfContentByte.Stroke();
        }

        private void CreatePDFFooter(PdfContentByte pdfContentByte)
        {
            pdfContentByte.SetLineWidth(0.5f);
            pdfContentByte.MoveTo(30f, 15f);
            pdfContentByte.LineTo(PageSize.A4.Width - 30f, 15f);
            int width = (int)(PageSize.A4.Width - 60f) / 3;
            pdfContentByte.MoveTo(30f, 30f);
            pdfContentByte.LineTo(PageSize.A4.Width - 30f, 30f);
            pdfContentByte.Stroke();
            Font font = this.GetFont("");
            pdfContentByte.SetColorFill(Color.GRAY);
            pdfContentByte.SetFontAndSize(font.BaseFont, 10f);
            int num = 34;
            int num1 = 20;
            string str = string.Concat("检验者：", this._tester);
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            num = 34;
            num1 = 5;
            str = string.Concat("日期：", this._testDateTime);
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            num = 34 + width;
            num1 = 20;
            str = string.Concat("校对者：", this._corrector);
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            num = 34 + width;
            num1 = 5;
            str = string.Concat("日期：", this._correctDateTime);
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            num = 34 + width * 2;
            num1 = 20;
            str = string.Concat("审核人：", this._auditor);
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            num = 34 + width * 2;
            num1 = 5;
            str = string.Concat("日期：", this._auditDateTime);
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            pdfContentByte.Stroke();
        }

        private void CreatePDFNext(PdfContentByte pdfContentByte)
        {
            Font font = this.GetFont("");
            pdfContentByte.SetColorFill(Color.GRAY);
            pdfContentByte.SetFontAndSize(font.BaseFont, 7f);
            int num = 5;
            int num1 = 5;
            string str = "(转下页)";
            pdfContentByte.SetTextMatrix((float)num, (float)num1);
            pdfContentByte.ShowText(str);
            pdfContentByte.Stroke();
        }

        private void CreatePDFPrevious(PdfContentByte pdfContentByte)
        {
            Font font = this.GetFont("");
            pdfContentByte.SetColorFill(Color.GRAY);
            pdfContentByte.SetFontAndSize(font.BaseFont, 7f);
            int num = 5;
            int height = (int)PageSize.A4.Height + 25;
            string str = "(接上页)";
            pdfContentByte.SetTextMatrix((float)num, (float)height);
            pdfContentByte.ShowText(str);
            pdfContentByte.Stroke();
        }

        public bool DelHeaderFooter(string strPDFPath, string strNewPDFPath)
        {
            bool flag = false;
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(strNewPDFPath, FileMode.Create));
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    PdfReader pdfReader = new PdfReader(strPDFPath);
                    int numberOfPages = pdfReader.NumberOfPages;
                    for (int i = 1; i <= numberOfPages; i++)
                    {
                        document.NewPage();
                        if ((int)pdfReader.GetPageContent(i).Length > 0)
                        {
                            directContent.AddTemplate(instance.GetImportedPage(pdfReader, i), 0f, 0f);
                        }
                    }
                    directContent.BeginText();
                    this.CreatePDFBLANKFooter(directContent);
                    directContent.EndText();
                    flag = true;
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
            finally
            {
                if (document.IsOpen())
                {
                    document.Close();
                }
            }
            return flag;
        }

        private Font GetFont(string strFont)
        {
            Font font;
            FontFactory.Register("C:\\Windows\\Fonts\\simsun.ttc");
            Font font1 = FontFactory.GetFont("宋体", "Identity-H", true, 10f);
            if ((strFont == null ? false : strFont.Length != 0))
            {
                string[] strArrays = strFont.Split(new char[] { ',' });
                font1.Size = float.Parse(strArrays[1].Trim().Substring(0, strArrays[1].Length - 3));
                for (int i = 2; i < (int)strArrays.Length; i++)
                {
                    string upper = strArrays[i].Trim().ToUpper();
                    if (upper != null)
                    {
                        if (upper == "BOLD")
                        {
                            font1.SetStyle(1);
                        }
                        else if (upper == "ITALIC")
                        {
                            font1.SetStyle(2);
                        }
                        else if (upper == "UNDERLINE")
                        {
                            font1.SetStyle(4);
                        }
                        else if (upper == "STRIKEOUT")
                        {
                            font1.SetStyle(8);
                        }
                    }
                }
                font = font1;
            }
            else
            {
                font = font1;
            }
            return font;
        }
    }
}
