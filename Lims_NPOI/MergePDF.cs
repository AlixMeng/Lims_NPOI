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
            string str = string.Concat("共 ", pagenum.ToString(), " 页 第 ", nownum.ToString(), " 页");
            pdfContentByte.SetTextMatrix((float)width, (float)height);
            pdfContentByte.ShowText(str);
            //width = (int)PageSize.A4.Width - 90;
            //height = (int)PageSize.A4.Height;
            //str = string.Concat("共 ", pagenum.ToString(), " 页,");
            //pdfContentByte.SetTextMatrix((float)width, (float)height);
            //pdfContentByte.ShowText(str);
            pdfContentByte.Stroke();
        }

        //将一个pdf的某一页插入另一个pdf的指定位置,位置索引从1开始
        /// <summary>
        /// 将一个pdf的某一页插入另一个pdf的指定位置,位置索引从1开始
        /// </summary>
        /// <param name="oPdf">被插入的pdf</param>
        /// <param name="oIndex">插入位置,从1开始</param>
        /// <param name="iPdf">要插入的pdf</param>
        /// <param name="iIndex">要插入的页索引</param>
        /// <param name="outPdf">输出pdf路径</param>
        /// <returns></returns>
        public bool InsertPageToPdf(string oPdf, int oIndex, string iPdf, int iIndex, string outPdf)
        {

            bool flag = false;
            if (outPdf.Length == 0)
            {
                flag = false;
                throw new Exception("合并后的文件名不能为空。");
            }
            if (!File.Exists(oPdf))
            {
                flag = false;
                throw new Exception("被插入pdf文件不存在。");
            }
            if (!File.Exists(iPdf))
            {
                flag = false;
                throw new Exception("要插入pdf的文件不存在。");
            }
            if (File.Exists(outPdf))
            {
                File.Delete(outPdf);
            }

            //设置页面大小
            Rectangle rectangle = new Rectangle(620.25f, 876.75f);
            //创建pdf文档大小
            Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
            try
            {
                try
                {
                    PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(outPdf, FileMode.Create));
                    document.Open();
                    PdfContentByte directContent = instance.DirectContent;
                    
                    //计算被插入的pdf页数
                    PdfReader oReader = new PdfReader(oPdf);
                    int oNumberOfPages = oReader.NumberOfPages;
                    for (int i = 1; i < oNumberOfPages + 2; i++)
                    {
                        //如果不等,则直接复制oPdf的当前页
                        if(i != oIndex)
                        {
                            document.NewPage();
                            if ((int)oReader.GetPageContent(i).Length > 0)
                            {
                                directContent.AddTemplate(instance.GetImportedPage(oReader, i), 0f, 0f);
                            }
                        }
                        //如果相等,则将要插入的pdf也插入指定位置
                        else
                        {
                            PdfReader iReader = new PdfReader(iPdf);
                            if ((int)iReader.GetPageContent(iIndex).Length > 0)
                            {
                                document.NewPage();
                                directContent.AddTemplate(instance.GetImportedPage(iReader, iIndex), 0f, 0f);
                                //break;
                                //插入pdf页后,改变索引以跳过此页
                                oNumberOfPages--;
                                oIndex--;
                                i--;                                
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
                if (document != null)
                {
                    if (document.IsOpen())
                    {
                        document.Close();
                    }
                    document = null;
                }
            }
            return flag;

        }

        /// <summary>
        /// 指定长宽的签名
        /// </summary>
        /// <param name="inputfilepath"></param>
        /// <param name="outputfilepath"></param>
        /// <param name="imgPath"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
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
                classLims_NPOI.WriteLog(ex, "");
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
        /// 资质章签名,固定在封面上方靠中
        /// </summary>
        /// <param name="inputfilepath"></param>
        /// <param name="outputfilepath"></param>
        /// <param name="imgPath"></param>
        /// <param name="x">图片中心: X坐标</param>
        /// <param name="y">图片中心:Y坐标</param>
        /// <param name="XYSign">无实际用途,用于方法的重载</param>
        public void addImageToPdf(string inputfilepath, string outputfilepath, string imgPath, double x, double y, string XYSign)
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
                if (img.DpiX != img.DpiY)
                {
                    classLims_NPOI.WriteLog("图片文件横向纵向分辨率不同,无法处理", "");
                    return;
                }
                var Alignment = Image.ALIGN_LEFT;
                img.Alignment = Alignment;
                //300 dpi的图片需要缩放到72 dpi对应的尺寸, 缩放倍数为24%
                float scalePct = 72f / img.DpiX;
                img.ScalePercent(scalePct * 100);
                //img.Height = img.Height * 0.24f;
                float imgWidth = img.Width * scalePct;
                float imgHeight = img.Height * scalePct;
                img.SetAbsolutePosition((float)x - (imgWidth / 2), (float)y - (imgHeight / 2));//设置图片坐标

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
                classLims_NPOI.WriteLog(ex, "");
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
        /// 资质章签名,固定在封面上方靠左
        /// </summary>
        /// <param name="inputfilepath"></param>
        /// <param name="outputfilepath"></param>
        /// <param name="imgPath"></param>
        /// <param name="x">图片中心: X坐标</param>
        /// <param name="y">图片中心:Y坐标</param>
        /// <param name="scale">缩放大小, 不缩放为100</param>
        public void addImageToPdf_Left(string inputfilepath, string outputfilepath, string imgPath, double x, double y, float scale)
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
                if (img.DpiX != img.DpiY)
                {
                    classLims_NPOI.WriteLog("图片文件横向纵向分辨率不同,无法处理", "");
                    return;
                }
                var Alignment = Image.ALIGN_LEFT;
                img.Alignment = Alignment;
                //300 dpi的图片需要缩放到72 dpi对应的尺寸, 缩放倍数为24%
                float scalePct = 72f / img.DpiX;
                img.ScalePercent(scalePct * scale);
                //img.Height = img.Height * 0.24f;
                float imgWidth = img.Width * scalePct;
                float imgHeight = img.Height * scalePct;
                img.SetAbsolutePosition((float)x, (float)y);//设置图片坐标,靠左

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
                classLims_NPOI.WriteLog(ex, "");
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

        //添加图片到pdf,指定坐标和统一的长宽
        /// <summary>
        /// 添加图片到pdf,指定坐标和统一的长宽
        /// </summary>
        /// <param name="inputfilepath">输入pdf路径</param>
        /// <param name="outputfilepath">输出pdf路径</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="x">X坐标</param>
        /// <param name="y">Y坐标</param>
        /// <param name="widthAndHeight">长宽数值,按像素</param>
        public void addImageToPdf(string inputfilepath, string outputfilepath, string imgPath, double x, double y, double widthAndHeight)
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
                if (img.DpiX != img.DpiY)
                {
                    classLims_NPOI.WriteLog("图片文件横向纵向分辨率不同,无法处理", "");
                    return;
                }
                if (img.Width != img.Height)
                {
                    classLims_NPOI.WriteLog("图片文件横向纵向像素数值不同,不适用于添加二维码", "");
                    return;
                }
                var Alignment = Image.ALIGN_LEFT;
                img.Alignment = Alignment;
                //设置缩放尺寸
                float scalePct = (float)widthAndHeight / img.Width;
                img.ScalePercent(scalePct * 100);
                //img.Height = img.Height * 0.24f;
                float width = img.Width * scalePct;
                float height = img.Height * scalePct;
                img.SetAbsolutePosition((float)x, (float)y);//设置图片坐标

                PdfContentByte content;
                PdfGState gs = new PdfGState();
                content = pdfStamper.GetOverContent(1);//在内容上方加水印,起始索引为1
                float f1 = content.XTLM;
                float f2 = content.WordSpacing;
                gs.FillOpacity = 1;//透明度,0为透明,1为完全不透明
                content.SetGState(gs);
                content.AddImage(img);
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
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

        //添加图片到pdf,指定坐标和长宽
        /// <summary>
        /// 添加图片到pdf,指定坐标和长宽
        /// </summary>
        /// <param name="inputfilepath">输入pdf路径</param>
        /// <param name="outputfilepath">输出pdf路径</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="x">X坐标</param>
        /// <param name="y">Y坐标</param>
        /// <param name="dWidth">长数值,按像素</param>
        /// <param name="dHeight">无用的图片高度,因为图片按宽度缩放并保留纵横比</param>
        public void addImageToPdf(string inputfilepath, string outputfilepath, string imgPath, double x, double y, double dWidth, string dHeight)
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

                var Alignment = Image.ALIGN_LEFT;
                img.Alignment = Alignment;
                //设置缩放尺寸, 按照宽度计算百分比
                float scalePct = (float)dWidth / img.Width;
                img.ScalePercent(scalePct * 100);
                //img.Height = img.Height * 0.24f;
                float width = img.Width * scalePct;
                float height = img.Height * scalePct;
                img.SetAbsolutePosition((float)x, (float)y);//设置图片坐标

                PdfContentByte content;
                PdfGState gs = new PdfGState();
                content = pdfStamper.GetOverContent(1);//在内容上方加水印,起始索引为1
                float f1 = content.XTLM;
                float f2 = content.WordSpacing;
                gs.FillOpacity = 1;//透明度,0为透明,1为完全不透明
                content.SetGState(gs);
                content.AddImage(img);
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
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
                    if (i == 1)
                    {
                        continue;
                    }
                    else if (i == 2)
                    {
                        if (SY_X < 0 || SY_Y < 0)
                        {
                            width = (int)PageSize.A4.Width - 40;
                            height = (int)PageSize.A4.Height - 128;
                        }
                        else
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

                    string waterMarkName = string.Concat("共 ", (total - 1).ToString(), " 页 第 ", (i - 1).ToString(), " 页");
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

        /// <summary>
        /// 添加普通偏转角度文字水印,指定了首页位置, 起始页码
        /// </summary>
        /// <param name="inputfilepath">输入pdf路径</param>
        /// <param name="outputfilepath">输出pdf路径</param>
        /// <param name="SY_X">首页页码x坐标,小于0时使用默认</param>
        /// <param name="SY_Y">首页页码y坐标,小于0时使用默认</param>
        /// <param name="FY_X">附页页码x坐标,小于0时使用默认</param>
        /// <param name="FY_Y">附页页码y坐标,小于0时使用默认</param>
        /// <param name="syIndex">首页索引</param>
        public void setPagesWatermark(string inputfilepath, string outputfilepath, float SY_X, float SY_Y, float FY_X, float FY_Y,
            int syIndex)
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
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont("C:\\WINDOWS\\Fonts\\simsun.ttc,1", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();
                //首页x=40, y=125
                //附页x=40, y=103
                for (int i = syIndex; i <= total; i++)
                {
                    #region 处理页数和页码坐标
                    if (i< syIndex)
                    {
                        continue;
                    }
                    else if (i == syIndex)
                    {
                        if (SY_X < 0 || SY_Y < 0)
                        {
                            width = (int)PageSize.A4.Width - 40;
                            height = (int)PageSize.A4.Height - 128;
                        }
                        else
                        {
                            width = SY_X;
                            height = SY_Y;
                        }

                    }
                    else if (i > syIndex)
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

                    string waterMarkName = string.Concat("共 ", (total - syIndex + 1).ToString(), " 页 第 ", (i - syIndex + 1).ToString(), " 页");
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

        /// <summary>
        /// 添加页码到指定页
        /// </summary>
        /// <param name="inputfilepath">输入pdf路径</param>
        /// <param name="outputfilepath">输出pdf路径</param>
        /// <param name="X">首页页码x坐标,小于0时使用默认</param>
        /// <param name="Y">首页页码y坐标,小于0时使用默认</param>
        /// <param name="pageIndex">页码编号</param>
        public void addPagenoToOnePage(string inputfilepath, string outputfilepath, float X, float Y, int pageIndex)
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
                    if (i != pageIndex)
                    {
                        continue;
                    }
                    else if (i == pageIndex)
                    {
                        if (X < 0 || Y < 0)
                        {
                            width = (int)PageSize.A4.Width - 40;
                            height = (int)PageSize.A4.Height - 128;
                        }
                        else
                        {
                            width = X;
                            height = Y;
                        }

                    }

                    #endregion

                    string waterMarkName = string.Concat("共 ", (total - pageIndex + 1).ToString(), " 页 第 ", (i - pageIndex + 1).ToString(), " 页");
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

        /// <summary>
        /// 添加页码到指定页
        /// </summary>
        /// <param name="inputfilepath">输入pdf路径</param>
        /// <param name="outputfilepath">输出pdf路径</param>
        /// <param name="X">首页页码x坐标,小于0时使用默认</param>
        /// <param name="Y">首页页码y坐标,小于0时使用默认</param>
        /// <param name="pageIndex">页码编号</param>
        /// <param name="waterString">水印值</param>
        public void addPagenoToSy(string inputfilepath, string outputfilepath, float X, float Y, int pageIndex, string waterString)
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
                    if (i != pageIndex)
                    {
                        continue;
                    }
                    else if (i == pageIndex)
                    {
                        if (X < 0 || Y < 0)
                        {
                            width = (int)PageSize.A4.Width - 40;
                            height = (int)PageSize.A4.Height - 128;
                        }
                        else
                        {
                            width = X;
                            height = Y;
                        }

                    }

                    #endregion

                    string waterMarkName = waterString;
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


        public Font GetFont(string strFont)
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

        //获取pdf指定页码的新pdf, 默认是第三页
        public string GetOnepage(string oldPdf, int pageIndex, string newPage)
        {
            try
            {
                if (!(File.Exists(oldPdf)))
                {
                    throw new Exception("要提取的文件不存在。");
                }
                if (File.Exists(newPage))
                {
                    File.Delete(newPage);
                }
                //创建新的pdf
                Rectangle rectangle = new Rectangle(620.25f, 876.75f);
                Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
                PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(newPage, FileMode.Create));

                //加载pdf并遍历页
                PdfReader pdfReader = new PdfReader(oldPdf);
                int numberOfPages = pdfReader.NumberOfPages;
                document.Open();
                PdfContentByte directContent = instance.DirectContent;
                for (int j = 1; j <= numberOfPages; j++)
                {
                    
                    if (j==pageIndex && (int)pdfReader.GetPageContent(j).Length > 0)
                    {
                        document.NewPage();
                        directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 0f, 0f);
                        break;
                    }
                }
                document.Close();
                return newPage;
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                Console.Error.WriteLine(exception.Message);
                Console.Error.WriteLine(exception.StackTrace);
            }
            return "";
        }

        //替换pdf的其中一页
        public string replaceOnePage(string oldPdf, int oldIndex, string onePage, string newPdf)
        {
            try
            {
                if (!(File.Exists(oldPdf)))
                {
                    throw new Exception("要处理的文件不存在。");
                }
                if (!(File.Exists(onePage)))
                {
                    throw new Exception("要插入的pdf不存在。");
                }
                if (File.Exists(newPdf))
                {
                    File.Delete(newPdf);
                }
                //创建新的pdf
                Rectangle rectangle = new Rectangle(620.25f, 876.75f);
                Document document = new Document(rectangle, 5f, 5f, 10f, 10f);
                PdfWriter instance = PdfWriter.GetInstance(document, new FileStream(newPdf, FileMode.Create));

                //加载pdf并遍历页
                PdfReader pdfReader = new PdfReader(oldPdf);
                int numberOfPages = pdfReader.NumberOfPages;
                document.Open();
                PdfContentByte directContent = instance.DirectContent;
                for (int j = 1; j <= numberOfPages; j++)
                {

                    if (j != oldIndex && (int)pdfReader.GetPageContent(j).Length > 0)
                    {
                        document.NewPage();
                        directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), 0f, 0f);
                    }
                    else if (j == oldIndex && (int)pdfReader.GetPageContent(j).Length > 0)
                    {
                        //加载新的pdf并提取第一页
                        PdfReader onePdfPage = new PdfReader(onePage);
                        document.NewPage();
                        directContent.AddTemplate(instance.GetImportedPage(onePdfPage, 1), 0f, 0f);
                    }
                }
                document.Close();
                return newPdf;
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                Console.Error.WriteLine(exception.Message);
                Console.Error.WriteLine(exception.StackTrace);
            }
            return "";
        }

        public int GetPageCount(string[] fileList)
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

        public int GetPageCount(string fileList)
        {
            int numberOfPages = 0;
            try
            {
                Document document = new Document();
                document.Open();
                PdfReader pdfReader = new PdfReader(fileList);
                numberOfPages = numberOfPages + pdfReader.NumberOfPages;
                
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

        #region 使用spire可方便地合并pdf,但是会收费
        ////Spire合并PDF
        //public bool mergePdfs(string[] fileList, string toPdfPath)
        //{
        //    PdfDocumentBase doc = Spire.Pdf.PdfDocument.MergeFiles(fileList);
        //    doc.Save(toPdfPath, FileFormat.PDF);
        //    System.Diagnostics.Process.Start(toPdfPath);
        //    return true;
        //}
        #endregion

        //不带缩放的合并pdf
        /// <summary>
        /// 不带缩放的合并pdf
        /// </summary>
        /// <param name="strFileList"></param>
        /// <param name="strMergeFile"></param>
        /// <returns></returns>
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
                    classLims_NPOI.WriteLog(exception, "");
                }
            }
            finally
            {
            }
            return flag;
        }

        /// <summary>
        /// 带缩放的合并pdf
        /// </summary>
        /// <param name="strFileList"></param>
        /// <param name="strMergeFile"></param>
        /// <param name="precent"></param>
        /// <returns></returns>
        public bool MergeAttachments(string strFileList, string strMergeFile, double precent)
        {
            bool flag = false;
            this._fileListToMerge = strFileList;
            this._mergeFile = strMergeFile;
            try
            {
                try
                {
                    flag = this.MergeAttachments(precent);
                }
                catch (Exception exception)
                {
                    classLims_NPOI.WriteLog(exception, "");
                }
            }
            finally
            {
            }
            return flag;
        }

        /// <summary>
        /// 不带缩放的合并pdf
        /// </summary>
        /// <returns></returns>
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
            //Rectangle rectangle = new Rectangle(0,0, 2480.671844f, 3505.521474f);
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

        /// <summary>
        /// 带缩放的合并
        /// </summary>
        /// <returns></returns>
        public bool MergeAttachments(double precent)
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
            //Rectangle rectangle = new Rectangle(0,0, 2480.671844f, 3505.521474f);
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
                                    directContent.AddTemplate(instance.GetImportedPage(pdfReader, j), (float)precent, 0f, 0f, (float)precent, 0f, 0f);
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


        //缩放pdf页面
        /// <summary>
        /// 缩放pdf页面
        /// </summary>
        /// <param name="docPath"></param>
        /// <param name="precent"></param>
        /// <param name="newPath"></param>
        public void setPdfPageSize(string docPath, float precent, string newPath)
        {
            try
            {
                if (File.Exists(newPath))
                {
                    File.Delete(newPath);
                }
                if (docPath.Equals(newPath))
                {
                    classLims_NPOI.WriteLog("缩放pdf页面时不能使用相同的输入输出文件路径!", "");
                    return;
                }

                PdfReader pdfReader = new PdfReader(docPath);
                Document document = new Document();
                PdfWriter pw = PdfWriter.GetInstance(document, new FileStream(newPath, FileMode.Create));
                document.Open();
                PdfContentByte directContent = pw.DirectContent;
                //页起始索引为1
                directContent.AddTemplate(pw.GetImportedPage(pdfReader, 1), precent, 0, 0, precent, 0, 0);
                document.Close();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
            return;
        }

        //图片设置为背景图,比例为√2:1
        public void setBackGroundPicture(string outputfilepath, string imgPath)
        {
            if (File.Exists(outputfilepath))
            {
                File.Delete(outputfilepath);
            }
            string tempFile = outputfilepath.Substring(0, outputfilepath.LastIndexOf("\\") + 2) + Guid.NewGuid().ToString() + ".pdf";
            //return;
            //第一步，创建一个 iTextSharp.text.Document对象的实例：
            Document document = new Document();
            PdfWriter pdfWriter = PdfWriter.GetInstance(document, new FileStream(tempFile, FileMode.Create));

            //第三步，打开当前Document
            document.Open();

            //第四步，为当前Document添加内容：
            //pdfWriter.a
            document.Add(new Paragraph("hello"));

            //第五步，关闭Document
            document.Close();
            //return;

            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(tempFile);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));

                int total = pdfReader.NumberOfPages;
                Image img = Image.GetInstance(imgPath);
                if (img.DpiX != img.DpiY)
                {
                    classLims_NPOI.WriteLog("图片文件横向纵向分辨率不同,无法处理", "");
                    return;
                }
                if (img.Height < img.Width)
                {
                    //计算旋转角度
                    img.Rotation = (float)Math.PI * 90 / 180;
                    img.ScaleAbsoluteHeight(620.25f);
                    img.ScaleAbsoluteWidth(876.75f);
                }
                else
                {
                    img.ScaleAbsoluteHeight(876.75f);
                    img.ScaleAbsoluteWidth(620.25f);
                }
                var Alignment = Image.ALIGN_LEFT;
                img.Alignment = Alignment;
                img.SetAbsolutePosition(0, 0);//设置图片坐标

                PdfContentByte content;
                PdfGState gs = new PdfGState();
                content = pdfStamper.GetOverContent(1);//在内容下方加水印,作为背景图,起始索引为1
                                                       //float f = content.GetEffectiveStringWidth("主检", true);
                float f1 = content.XTLM;
                float f2 = content.WordSpacing;
                gs.FillOpacity = 1;//透明度,0为透明,1为完全不透明
                content.SetGState(gs);
                content.AddImage(img);
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
            finally
            {
                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
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
