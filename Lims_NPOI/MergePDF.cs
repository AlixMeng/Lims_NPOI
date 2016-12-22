using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;

namespace nsLims_NPOI
{
    /// <summary>
    /// 合并pdf和pdf常用处理,使用iTextSharp类库
    /// </summary>
    class MergePDF
    {
        private string _fileListToMerge = "";

        private string _mergeFile = "";

        private int _MinusPageNum = 0;

        public string FileListToMerge
        {
            get
            {
                return this._fileListToMerge;
            }
            set
            {
                this._fileListToMerge = value;
            }
        }

        public string MergeFile
        {
            get
            {
                return this._mergeFile;
            }
            set
            {
                this._mergeFile = value;
            }
        }

        public int MinusPageNum
        {
            get
            {
                return this._MinusPageNum;
            }
            set
            {
                this._MinusPageNum = value;
            }
        }

        public MergePDF()
        {
        }

        private void CreateReportPageNum(PdfContentByte pdfContentByte, int pagenum, int nownum, int x, int y)
        {
            Font font = this.GetFont("");
            pdfContentByte.SetColorFill(Color.BLACK);
            pdfContentByte.SetFontAndSize(font.BaseFont, 10f);
            //附页x=125,y=103
            //首页x=   ,y=103
            int width = (int)PageSize.A4.Width - x;
            int height = (int)PageSize.A4.Height - y;
            string str = string.Concat("共 ", pagenum.ToString(), " 页,第 ", nownum.ToString(), " 页");
            pdfContentByte.SetTextMatrix((float)width, (float)height);
            pdfContentByte.ShowText(str);
            //width = (int)PageSize.A4.Width - 90;
            //height = (int)PageSize.A4.Height;
            //str = string.Concat("共 ", pagenum.ToString(), " 页,");
            //pdfContentByte.SetTextMatrix((float)width, (float)height);
            //pdfContentByte.ShowText(str);
            pdfContentByte.Stroke();
        }

        public void addImageToPdf(string inputfilepath, string outputfilepath, string imgPath, float width, float height)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            if (inputfilepath == outputfilepath)
            {
                classLims_NPOI.WriteLog("输入pdf文件和输出pdf文件不能相同", "");
                return;
            }
            if (File.Exists(outputfilepath))
            {
                File.Delete(outputfilepath);
            }
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));

                int total = pdfReader.NumberOfPages;
                Image img = Image.GetInstance(imgPath);
                img.SetAbsolutePosition(440f, 60f);//设置图片坐标,(0,0)为左下角
                //var Alignment = Image.ALIGN_LEFT;
                img.ScaleAbsolute(width, height);
                
                PdfContentByte content;                
                PdfGState gs = new PdfGState();
                content = pdfStamper.GetOverContent(1);//在内容上方加水印,起始索引为1
                //float f = content.GetEffectiveStringWidth("主检", true);
                float f1 = content.XTLM;
                float f2 = content.WordSpacing;
                gs.FillOpacity = 1;//透明度,0为透明,1为完全不透明
                content.SetGState(gs);
                content.AddImage(img);
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex,"");
                return;
            }
            finally
            {
                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }


        }

        /// <summary>
        /// 添加普通偏转角度文字水印
        /// </summary>
        /// <param name="inputfilepath">输入pdf路径</param>
        /// <param name="outputfilepath">输出pdf路径</param>
        /// <param name="SY_X">首页页码x坐标,小于0时使用默认</param>
        /// <param name="SY_Y">首页页码y坐标,小于0时使用默认</param>
        /// <param name="FY_X">附页页码x坐标,小于0时使用默认</param>
        /// <param name="FY_Y">附页页码y坐标,小于0时使用默认</param>
        public void setPagesWatermark(string inputfilepath, string outputfilepath, float SY_X, float SY_Y, float FY_X, float FY_Y)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            if(inputfilepath== outputfilepath)
            {
                classLims_NPOI.WriteLog("输入pdf文件和输出pdf文件不能相同", "");
                return;
            }
            if(File.Exists(outputfilepath))
            {
                File.Delete(outputfilepath);
            }
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                int total = pdfReader.NumberOfPages;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();
                //首页x=40, y=125
                //附页x=40, y=103
                for (int i = 1; i <= total; i++)
                {
                    #region 处理页数和页码坐标
                    if(i==1)
                    {
                        continue;
                    }
                    else if (i == 2)
                    {
                        if(SY_X<0 || SY_Y < 0)
                        {
                            width = (int)PageSize.A4.Width - 40;
                            height = (int)PageSize.A4.Height - 128;
                        }else
                        {
                            width = SY_X;
                            height = SY_Y;
                        }
                        
                    }
                    else if (i > 2)
                    {
                        if (FY_X < 0 || FY_Y < 0)
                        {
                            width = (int)PageSize.A4.Width - 54;
                            height = (int)PageSize.A4.Height - 113;
                        }
                        else
                        {
                            width = FY_X;
                            height = FY_Y;
                        }
                       
                    }
                    #endregion

                    string waterMarkName = string.Concat("共 ", (total-1).ToString(), " 页,第 ", (i-1).ToString(), " 页");
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    //透明度,0为透明,1为完全不透明
                    gs.FillOpacity = 1;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);
                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(Color.BLACK);
                    content.SetFontAndSize(font, 10f);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_RIGHT, waterMarkName, width, height, 0);
                    //content.SetColorFill(BaseColor.BLACK);
                    //content.SetFontAndSize(font, 8);
                    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
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

        private int GetPageCount(string[] fileList)
        {
            int numberOfPages = 0;
            try
            {
                Document document = new Document();
                document.Open();
                for (int i = 0; i < (int)fileList.Length; i++)
                {
                    PdfReader pdfReader = new PdfReader(fileList[i]);
                    numberOfPages = numberOfPages + pdfReader.NumberOfPages;
                }
                document.Close();
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                Console.Error.WriteLine(exception.Message);
                Console.Error.WriteLine(exception.StackTrace);
            }
            return numberOfPages;
        }

        public bool Merge(string strFileList, string strMergeFile)
        {
            bool flag = false;
            this._fileListToMerge = strFileList;
            this._mergeFile = strMergeFile;
            try
            {
                try
                {
                    flag = this.Merge();
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

        public void iTextSharpAddImage2Pdf(string filePath, string imgPath)
        {

            try
            {
                PdfReader reader = new PdfReader(filePath);     //读取现有的PDF文档
                iTextSharp.text.Rectangle psize = reader.GetPageSize(1);      //获取第一页
                Document document = new Document(psize, 50, 50, 50, 50);     //设置位置  
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Append));

                document.Open();

                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imgPath);           //插入图片

                img.SetAbsolutePosition(0, 0);

                writer.DirectContent.AddImage(img);         //添加图片
                document.Close();                
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }

        }


        public bool Merge()
        {
            bool flag = false;
            if (this._fileListToMerge.Length == 0)
            {
                flag = false;
                throw new Exception("要合并的文件列表不能为空。");
            }
            if (this._mergeFile.Length == 0)
            {
                flag = false;
                throw new Exception("合并后的文件名不能为空。");
            }
            if (File.Exists(this._mergeFile))
            {
                File.Delete(this._mergeFile);
            }
            string[] strArrays = this._fileListToMerge.Split(new char[] { ',' });
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(this._mergeFile, FileMode.Create));
                    //BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", "Identity-H", false);
                    BaseFont baseFont = BaseFont.CreateFont();
                    Font font = new Font(baseFont, 8f);
                    int pageCount = this.GetPageCount(strArrays);
                    HeaderFooter headerFooter = new HeaderFooter(new Phrase("第", font), new Phrase(string.Concat("页 共", pageCount.ToString(), "页"), font))
                    {
                        Border = 0
                    };
                    headerFooter.SetAlignment("Right");
                    document.Header = headerFooter;
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    for (int i = 0; i < (int)strArrays.Length; i++)
                    {
                        if (File.Exists(strArrays[i]))
                        {
                            PdfReader pdfReader = new PdfReader(strArrays[i]);
                            int numberOfPages = pdfReader.NumberOfPages;
                            for (int j = 1; j <= numberOfPages; j++)
                            {
                                document.NewPage();
                                if ((int)pdfReader.GetPageContent(j).Length > 0)
                                {
                                    directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 0f, 0f);
                                }
                            }
                        }
                    }
                    directContent.SetFontAndSize(baseFont, 4f);
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
                document = null;
            }
            return flag;
        }

        ////Spire合并PDF
        //public bool mergePdfs(string[] fileList, string toPdfPath)
        //{
        //    PdfDocumentBase doc = Spire.Pdf.PdfDocument.MergeFiles(fileList);
        //    doc.Save(toPdfPath, FileFormat.PDF);
        //    System.Diagnostics.Process.Start(toPdfPath);
        //    return true;
        //}

        public bool MergeAttachments(string strFileList, string strMergeFile)
        {
            bool flag = false;
            this._fileListToMerge = strFileList;
            this._mergeFile = strMergeFile;
            try
            {
                try
                {
                    flag = this.MergeAttachments();
                }
                catch (Exception exception)
                {
                    classLims_NPOI.WriteLog(exception,"");
                }
            }
            finally
            {
            }
            return flag;
        }

        public bool MergeAttachments()
        {
            bool flag = false;
            if (this._fileListToMerge.Length == 0)
            {
                flag = false;
                throw new Exception("要合并的文件列表不能为空。");
            }
            if (this._mergeFile.Length == 0)
            {
                flag = false;
                throw new Exception("合并后的文件名不能为空。");
            }
            if (File.Exists(this._mergeFile))
            {
                File.Delete(this._mergeFile);
            }
            string[] strArrays = this._fileListToMerge.Split(new char[] { ',' });
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(this._mergeFile, FileMode.Create));
                    BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", "Identity-H", false);
                    Font font = new Font(baseFont, 8f);
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    for (int i = 0; i < (int)strArrays.Length; i++)
                    {
                        if (File.Exists(strArrays[i]))
                        {
                            PdfReader pdfReader = new PdfReader(strArrays[i]);
                            int numberOfPages = pdfReader.NumberOfPages;
                            for (int j = 1; j <= numberOfPages; j++)
                            {
                                document.NewPage();
                                if ((int)pdfReader.GetPageContent(j).Length > 0)
                                {
                                    directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 0f, 0f);
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
                document = null;
            }
            return flag;
        }

        public bool MergeOne(string strFileList, string strMergeFile)
        {
            bool flag = false;
            this._fileListToMerge = strFileList;
            this._mergeFile = strMergeFile;
            try
            {
                try
                {
                    flag = this.MergeOne();
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

        public bool MergeOne()
        {
            bool flag = false;
            if (this._fileListToMerge.Length == 0)
            {
                flag = false;
                throw new Exception("要合并的文件列表不能为空。");
            }
            if (this._mergeFile.Length == 0)
            {
                flag = false;
                throw new Exception("合并后的文件名不能为空。");
            }
            if (File.Exists(this._mergeFile))
            {
                File.Delete(this._mergeFile);
            }
            string[] strArrays = this._fileListToMerge.Split(new char[] { ',' });
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(this._mergeFile, FileMode.Create));
                    BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", "Identity-H", false);
                    Font font = new Font(baseFont, 8f);
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    for (int i = 0; i < (int)strArrays.Length; i++)
                    {
                        if (File.Exists(strArrays[i]))
                        {
                            PdfReader pdfReader = new PdfReader(strArrays[i]);
                            int numberOfPages = pdfReader.NumberOfPages;
                            for (int j = 1; j <= numberOfPages; j++)
                            {
                                document.NewPage();
                                if ((int)pdfReader.GetPageContent(j).Length > 0)
                                {
                                    directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 0f, 0f);
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
                document = null;
            }
            return flag;
        }

        public bool MergeReport(string strFileList, string strMergeFile)
        {
            bool flag = false;
            this._fileListToMerge = strFileList;
            this._mergeFile = strMergeFile;
            try
            {
                try
                {
                    flag = this.MergeReport();
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

        public bool MergeReport()
        {
            bool flag = false;
            if (this._fileListToMerge.Length == 0)
            {
                flag = false;
                throw new Exception("要合并的文件列表不能为空。");
            }
            if (this._mergeFile.Length == 0)
            {
                flag = false;
                throw new Exception("合并后的文件名不能为空。");
            }
            if (File.Exists(this._mergeFile))
            {
                File.Delete(this._mergeFile);
            }
            string[] strArrays = this._fileListToMerge.Split(new char[] { ',' });
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(this._mergeFile, FileMode.Create));
                    BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", "Identity-H", false);
                    Font font = new Font(baseFont, 8f);
                    int pageCount = this.GetPageCount(strArrays);
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    int num = 0;
                    for (int i = 0; i < (int)strArrays.Length; i++)
                    {
                        if (File.Exists(strArrays[i]))
                        {
                            PdfReader pdfReader = new PdfReader(strArrays[i]);
                            int numberOfPages = pdfReader.NumberOfPages;
                            for (int j = 1; j <= numberOfPages; j++)
                            {
                                num++;
                                document.NewPage();
                                if ((int)pdfReader.GetPageContent(j).Length > 0)
                                {
                                    directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 10f, 10f);
                                }
                                directContent.BeginText();
                                this.CreateReportPageNum(directContent, pageCount, num, 125, 103);
                                directContent.EndText();
                            }
                        }
                    }
                    directContent.SetFontAndSize(baseFont, 4f);
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
                document = null;
            }
            return flag;
        }

        public bool MergeTwo(string strFileList, string strMergeFile, int MinusPageNum)
        {
            bool flag = false;
            this._fileListToMerge = strFileList;
            this._mergeFile = strMergeFile;
            this._MinusPageNum = MinusPageNum;
            try
            {
                try
                {
                    flag = this.MergeTwo();
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

        public bool MergeTwo()
        {
            bool flag = false;
            if (this._fileListToMerge.Length == 0)
            {
                flag = false;
                throw new Exception("要合并的文件列表不能为空。");
            }
            if (this._mergeFile.Length == 0)
            {
                flag = false;
                throw new Exception("合并后的文件名不能为空。");
            }
            if (File.Exists(this._mergeFile))
            {
                File.Delete(this._mergeFile);
            }
            int num = this._MinusPageNum;
            if (num < 0)
            {
                num = 0;
            }
            string[] strArrays = this._fileListToMerge.Split(new char[] { ',' });
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(this._mergeFile, FileMode.Create));
                    BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", "Identity-H", false);
                    Font font = new Font(baseFont, 8f);
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    for (int i = 0; i < (int)strArrays.Length; i++)
                    {
                        if (File.Exists(strArrays[i]))
                        {
                            PdfReader pdfReader = new PdfReader(strArrays[i]);
                            int numberOfPages = pdfReader.NumberOfPages;
                            if (num >= numberOfPages)
                            {
                                num = 0;
                            }
                            for (int j = 1; j <= numberOfPages - num; j++)
                            {
                                document.NewPage();
                                if ((int)pdfReader.GetPageContent(j).Length > 0)
                                {
                                    directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 0f, 0f);
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
                document = null;
            }
            return flag;
        }
    }
}
