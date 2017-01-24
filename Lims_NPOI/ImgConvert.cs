using System;
using System.Drawing;
using System.IO;
using Spire.Pdf;
using Spire.Pdf.General.Find;
using Spire.Pdf.Graphics;

namespace nsLims_NPOI
{
    /// <summary>
    /// 使用Spire.pdf-free操作pdf,主要使用添加图片功能
    /// </summary>
    class ImgConvert
    {
        public ImgConvert()
        {
        }

        #region unused code

        public byte[] load_pictMemory(string filePath)
        {
            byte[] numArray = null;
            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Exists)
            {
                numArray = new byte[checked(fileInfo.Length)];
                FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);
                BinaryReader binaryReader = new BinaryReader(fileStream);
                binaryReader.Read(numArray, 0, Convert.ToInt32(fileInfo.Length));
                fileStream.Dispose();
            }
            return numArray;
        }

        public void update_picture(string filePath, string filePathNew)
        {
            byte[] numArray = this.load_pictMemory(filePath);
            Image image = Image.FromStream(new MemoryStream(numArray));
            int width = image.Width;
            int height = image.Height;
            float horizontalResolution = image.HorizontalResolution;
            float verticalResolution = image.VerticalResolution;
            float single = (float)height / verticalResolution;
            float single1 = (float)width / horizontalResolution;
            double num = (double)single / 0.39;
            int num1 = (int)((double)height / num);
            int num2 = (int)((double)width / num);
            (new Bitmap(image, num2, num1)).Save(filePathNew);
        }
        #endregion

        //向pdf指定位置添加图片,按照坐标
        /// <summary>
        /// 向PDF指定位置添加图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="newPath"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="BlobFiled">图片字符串</param>
        /// <param name="imgHeight">图片高度,为0时使用图片原始高度</param>
        /// <param name="imgWidth">图片宽度,为0时使用图片原始宽度</param>
        public void addImage2Pdf_point(string filePath, string newPath, float X, float Y, string BlobFiled, float imgWidth, float imgHeight)
        {

            try
            {

                //Create a pdf document
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(filePath);

                //get the page
                PdfPageBase page = doc.Pages[0];

                byte[] numArray = System.Text.Encoding.UTF8.GetBytes(BlobFiled);

                //get the image
                if (numArray == null)
                {
                    return;
                }
                PdfImage image = PdfImage.FromStream(new MemoryStream(numArray));
                float width = imgWidth;
                float height = imgHeight;
                if (imgWidth == 0)
                {
                    width = image.Width;
                }
                if (imgHeight == 0)
                {
                    height = image.Height;
                }

                //insert image
                page.Canvas.DrawImage(image, X, Y, width, height);

                //save pdf file
                doc.SaveToFile(newPath);
                doc.Close();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }

        }

        /// <summary>
        /// 向pdf添加背景图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="newPath"></param>
        /// <param name="imgPath"></param>
        public void addBackgroundImg2Pdf(string filePath, string newPath, string imgPath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return;
                }
                //Create a pdf document
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(filePath);


                //get the page
                PdfPageBase page = doc.Pages[0];

                if (!File.Exists(imgPath))
                {
                    return;
                }
                //插入一个背景图片
                System.Drawing.Image img = System.Drawing.Image.FromFile(imgPath);
                page.BackgroundImage = img;

                //save pdf file
                doc.SaveToFile(newPath);
                doc.Close();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
        }

        
        //绘制文本到pdf
        private static void AlignText(PdfPageBase page, string label, float x, float y)
        {

            //使用仿宋字体才不会报错
            PdfTrueTypeFont font1 = new PdfTrueTypeFont(@"C:\Windows\Fonts\simfang.ttf", 10f);
            PdfSolidBrush brush = new PdfSolidBrush(System.Drawing.Color.Black);

            PdfStringFormat rightAlignment = new PdfStringFormat(PdfTextAlignment.Right, PdfVerticalAlignment.Middle);
            page.Canvas.DrawString(label, font1, brush, x, y, rightAlignment);

        }


        //pdf写入字符串
        public static void ReplaceFlag(string filePath, string newPath, int startPageIndex)
        {

            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(filePath);
            //起始页码超过总页码
            if (startPageIndex >= doc.Pages.Count)
            {
                return;
            }
            //apply template in PDF document
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                PdfPageBase page = doc.Pages[i];
                float x, y;
                if (i < startPageIndex)//封面
                    continue;
                else if (i == startPageIndex)//首页
                {
                    string label = "共 " + (doc.Pages.Count - startPageIndex).ToString() + " 页,第 1 页";
                    x = page.Canvas.ClientSize.Width - 40;
                    y = 120.8f;
                    AlignText(page, label, x, y);
                }
                else if (i > 1)//附页
                {
                    string label = "共 " + (doc.Pages.Count - startPageIndex).ToString() + " 页,第 " + (i - startPageIndex

+ 1).ToString() + " 页";
                    AlignText(page, label, page.Canvas.ClientSize.Width - 40, 110f);
                }
            }


            doc.SaveToFile(newPath);
        }

        //向签发日期后添加日期
        /// <summary>
        /// 向签发日期后添加日期
        /// </summary>
        /// <param name="oldPath"></param>
        /// <param name="newPath"></param>
        /// <param name="flagTxt">要查找的字符串</param>
        /// <param name="issueDate">要写入的日期</param>
        public void addIssueDateToPdf(string oldPath, string newPath, string flagTxt, string issueDate)
        {
            try
            {
                //Create a pdf document
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(oldPath);

                PdfPageBase page = null;
                PdfTextFind[] result = null;
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    page = doc.Pages[i];
                    result = page.FindText(flagTxt).Finds;
                    if (result.Length > 0)
                    {
                        break;
                    }
                }
                //如果没找到标记字符串,记录并返回
                if (result == null || result.Length == 0)
                {
                    classLims_NPOI.WriteLog("标记字符串未找到", "");
                    return;
                }

                //获取第一次出现文字的坐标，宽度和高度  
                PointF pointf = result[0].Position;
                //获取文字的宽高
                SizeF size = result[0].Size;

                AlignText(page, issueDate, pointf.X + size.Width  + 50, pointf.Y + 3);
                //save pdf file
                doc.SaveToFile(newPath);
                doc.Close();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }
        }

        //向PDF指定位置添加图片
        /// <summary>
        /// 向PDF指定位置添加图片,按照标记字符串
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="newPath"></param>
        /// <param name="flagTxt">标记字符串</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="imgHeight">图片高度,为0时使用图片原始高度</param>
        /// <param name="imgWidth">图片宽度,为0时使用图片原始宽度</param>
        public void addImage2Pdf_path(string filePath, string newPath, string flagTxt, string imgPath, float imgWidth, float imgHeight)
        {

            try
            {
                //Create a pdf document
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(filePath);

                PdfPageBase page = null;
                PdfTextFind[] result = null;
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    page = doc.Pages[i];
                    result = page.FindText(flagTxt).Finds;
                    if (result.Length > 0)
                    {
                        break;
                    }
                }
                //如果没找到标记字符串,记录并返回
                if (result == null || result.Length == 0)
                {
                    classLims_NPOI.WriteLog("标记字符串未找到", "");
                    return;
                }

                //获取第一次出现文字的坐标，宽度和高度  
                PointF pointf = result[0].Position;
                //获取文字的宽高
                SizeF size = result[0].Size;

                if (!File.Exists(imgPath))
                {
                    classLims_NPOI.WriteLog("图片文件未找到", "");
                    return;
                }
                PdfImage image = PdfImage.FromFile(imgPath);
                float width = imgWidth;
                float height = imgHeight;
                if (imgWidth == 0)
                {
                    width = image.Width;
                }
                if (imgHeight == 0)
                {
                    height = image.Height;
                }


                //签名图片Y轴位置为 始终和标记字符串居中对齐
                pointf.Y = pointf.Y - (height - size.Height) / 2;
                //insert image
                page.Canvas.DrawImage(image, pointf.X, pointf.Y, width, height);

                //save pdf file
                doc.SaveToFile(newPath);
                doc.Close();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }

        }

        /// <summary>
        /// 合并多个pdf
        /// </summary>
        /// <param name="fileList">要合并的pdf数组</param>
        /// <param name="toFile">目标文件路径</param>
        /// <returns></returns>
        public bool MergeBySpire(object[] fileList, string toFile)
        {
            try
            {
                string[] fList;
                fList = classLims_NPOI.dArray2String1(fileList);
                PdfDocumentBase doc = PdfDocument.MergeFiles(fList);
                doc.Save(toFile, FileFormat.PDF);
                return true;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return false;
            }
        }

        //向PDF指定位置添加图片
        /// <summary>
        /// 向PDF指定位置添加图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="newPath"></param>
        /// <param name="flagTxt">标记字符串</param>
        /// <param name="BlobFiled">图片字符串</param>
        /// <param name="imgHeight">图片高度,为0时使用图片原始高度</param>
        /// <param name="imgWidth">图片宽度,为0时使用图片原始宽度</param>
        public void addImage2Pdf(string filePath, string newPath, string flagTxt, string BlobFiled, float imgWidth, float imgHeight)
        {

            try
            {

                //Create a pdf document
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(filePath);


                //get the page
                PdfPageBase page = doc.Pages[0];

                //获取文本在PdfPage的坐标
                PointF pointf = new PointF();
                PdfTextFind[] result = page.FindText(flagTxt).Finds;


                //如果没找到标记字符串,直接返回
                if (result.Length == 0)
                {
                    classLims_NPOI.WriteLog("标记字符串未找到", "");
                    return;
                }
                //获取第一次出现文字的坐标，宽度和高度  
                pointf = result[0].Position;
                //获取文字的宽高
                SizeF size = result[0].Size;

                byte[] numArray = System.Text.Encoding.UTF8.GetBytes(BlobFiled);
                if (numArray == null)
                {
                    classLims_NPOI.WriteLog("字节数组为空", "");
                    return;
                }

                PdfImage image = PdfImage.FromStream(new MemoryStream(numArray));
                float width = imgWidth;
                float height = imgHeight;
                if (imgWidth == 0)
                {
                    width = image.Width;
                }
                if (imgHeight == 0)
                {
                    height = image.Height;
                }


                //签名图片Y轴位置为 始终和标记字符串居中对齐
                pointf.Y = pointf.Y - (height - size.Height) / 2;
                //insert image
                page.Canvas.DrawImage(image, pointf.X, pointf.Y, width, height);

                //save pdf file
                doc.SaveToFile(newPath);
                doc.Close();
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
            }

        }

        public void string2Image(string str, string imgPath)
        {
            Graphics g = Graphics.FromImage(new Bitmap(1, 1));
            Font font = new Font("宋体", 10);
            SizeF sizeF = g.MeasureString(str, font); //测量出字体的高度和宽度  
            Brush brush; //笔刷，颜色  
            brush = Brushes.Lime;
            PointF pf = new PointF(0, 0);
            Bitmap img = new Bitmap(Convert.ToInt32(sizeF.Width), Convert.ToInt32(sizeF.Height));
            g = Graphics.FromImage(img);
            g.DrawString(str, font, brush, pf);
            //输出图片  
            MemoryStream ms = new MemoryStream();
            img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            //Response.BinaryWrite(ms.ToArray());
        }
        
    }
}
