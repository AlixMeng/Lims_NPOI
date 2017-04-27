using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;
using System.IO;
using WORD = Microsoft.Office.Interop.Word;
//using DocumentFormat.OpenXml.Wordprocessing;

namespace nsLims_NPOI
{
    /// <summary>
    /// 操作word常用方法
    /// </summary>
    class DocXAction
    {
        public DocXAction()
        {
        }

        #region 操作word,用DocX

        /// <summary>
        /// 获取word页码数
        /// </summary>
        /// <param name="strSourceFile">要转换的Word文档</param>
        /// <returns>页码数</returns>
        public int getWordPages(object strSourceFile)
        {
            try
            {
                int pages = 0;
                if (File.Exists(strSourceFile.ToString()))
                {
                    object Nothing = System.Reflection.Missing.Value;
                    //创建一个名为WordApp的组件对象 
                    WORD.Application wordApp = new WORD.ApplicationClass();
                    //创建一个名为WordDoc的文档对象并打开
                    WORD.Document doc = wordApp.Documents.Open(ref strSourceFile, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    //下面是取得打开文件的页数  
                    pages = doc.ComputeStatistics(WORD.WdStatistic.wdStatisticPages, ref Nothing);
                    //关闭文档对象
                    object saveOption = WORD.WdSaveOptions.wdDoNotSaveChanges;
                    ((Microsoft.Office.Interop.Word._Document)doc).Close(ref saveOption, ref Nothing, ref Nothing);
                    //推出组建   
                    wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
                }
                return pages;
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return 0;
            }


        }

        /// <summary>
        /// 根据字符串获取字符串所在段落
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <returns></returns>
        public Paragraph GetParagraphByReplaceFlag(DocX doc, string replaceFlag, string alignment)
        {
            List<Paragraph> lstParagraphInHeaderFirst = null;
            List<Paragraph> lstParagraphInHeaderOdd = null;
            List<Paragraph> lstParagraphInHeaderEven = null;
            List<Paragraph> lstParagraphInFooterFirst = null;
            List<Paragraph> lstParagraphInFooterOdd = null;
            List<Paragraph> lstParagraphInFooterEven = null;

            if (doc.Headers.first != null)
            {
                lstParagraphInHeaderFirst = doc.Headers.first.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Headers.odd != null)
            {
                lstParagraphInHeaderOdd = doc.Headers.odd.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Headers.even != null)
            {
                lstParagraphInHeaderEven = doc.Headers.even.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Footers.first != null)
            {
                lstParagraphInFooterFirst = doc.Footers.first.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Footers.odd != null)
            {
                lstParagraphInFooterOdd = doc.Footers.odd.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Footers.even != null)
            {
                lstParagraphInFooterEven = doc.Footers.even.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }


            List<Paragraph> lstParagraph = doc.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();

            Paragraph p = null;
            Boolean bBreakOutOfFor = false;

            if (lstParagraphInHeaderFirst != null && lstParagraphInHeaderFirst.Count != 0)
            {
                p = lstParagraphInHeaderFirst[0];
            }
            else if (lstParagraphInHeaderOdd != null && lstParagraphInHeaderOdd.Count != 0)
            {
                p = lstParagraphInHeaderOdd[0];
            }
            else if (lstParagraphInHeaderEven != null && lstParagraphInHeaderEven.Count != 0)
            {
                p = lstParagraphInHeaderEven[0];
            }
            else if (lstParagraphInFooterFirst != null && lstParagraphInFooterFirst.Count != 0)
            {
                p = lstParagraphInFooterFirst[0];
            }
            else if (lstParagraphInFooterOdd != null && lstParagraphInFooterOdd.Count != 0)
            {
                p = lstParagraphInFooterOdd[0];
            }
            else if (lstParagraphInFooterEven != null && lstParagraphInFooterEven.Count != 0)
            {
                p = lstParagraphInFooterEven[0];
            }
            else if (lstParagraph.Count != 0)
            {
                p = lstParagraph[0];
            }
            else if (doc.Headers.first != null || doc.Headers.odd != null || doc.Headers.even != null ||
                    doc.Footers.first != null || doc.Footers.odd != null || doc.Footers.even != null)
            {
                List<Table> lstTables = null;
                if (doc.Headers.first != null && doc.Headers.first.Tables.Count != 0)
                {
                    lstTables = doc.Headers.first.Tables;
                }
                else if (doc.Headers.odd != null && doc.Headers.odd.Tables.Count != 0)
                {
                    lstTables = doc.Headers.odd.Tables;
                }
                else if (doc.Headers.even != null && doc.Headers.even.Tables.Count != 0)
                {
                    lstTables = doc.Headers.even.Tables;
                }
                else if (doc.Footers.first != null && doc.Footers.first.Tables.Count != 0)
                {
                    lstTables = doc.Footers.first.Tables;
                }
                else if (doc.Footers.odd != null && doc.Footers.odd.Tables.Count != 0)
                {
                    lstTables = doc.Footers.odd.Tables;
                }
                else if (doc.Footers.even != null && doc.Footers.even.Tables.Count != 0)
                {
                    lstTables = doc.Footers.even.Tables;
                }
                else
                {
                    lstTables = doc.Tables;
                }

                int nTablesInDoc = lstTables.Count;

                for (int i = 0; i < nTablesInDoc; i++)
                {
                    for (int m = 0; m < lstTables[i].RowCount; m++)
                    {
                        for (int n = 0; n < lstTables[i].Rows[m].Cells.Count; n++)
                        {
                            List<Paragraph> lstParagraphInCell = lstTables[i].Rows[m].Cells[n].Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();

                            if (lstParagraphInCell.Count != 0)
                            {
                                p = lstParagraphInCell[0];
                                bBreakOutOfFor = true;
                            }

                            if (bBreakOutOfFor)
                            {
                                break;
                            }

                            if ((m == lstTables[i].RowCount - 1) &&
                                (n == lstTables[i].Rows[m].Cells.Count - 1) &&
                                (lstParagraphInCell.Count == 0))
                            {
                                return null;
                            }

                        }

                        if (bBreakOutOfFor)
                        {
                            break;
                        }
                    }

                    if (bBreakOutOfFor)
                    {
                        break;
                    }
                }
            }

            if (alignment != null && alignment.Length > 0 && p != null)
            {
                if (alignment.ToUpper() != "LEFT" && alignment.ToUpper() != "RIGHT" && alignment.ToUpper() != "CENTER" && alignment.ToUpper() != "BOTH")
                {
                    p.Alignment = Alignment.left;
                }

                if (alignment.ToUpper() == "LEFT")
                {
                    p.Alignment = Alignment.left;
                }
                else if (alignment.ToUpper() == "RIGHT")
                {
                    p.Alignment = Alignment.right;
                }
                else if (alignment.ToUpper() == "CENTER")
                {
                    p.Alignment = Alignment.center;
                }
                else
                {
                    p.Alignment = Alignment.both;
                }
            }

            return p;
        }

        //获取带图片的段落
        /// <summary>
        /// 获取带图片的段落
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>段落数组</returns>
        private List<Paragraph> getPictureParagraphs(DocX doc)
        {
            List<Paragraph> listPar = new List<Paragraph>();
            foreach (Paragraph pointPar in doc.Paragraphs)
            {
                if (pointPar.Pictures.Count > 0)
                    listPar.Add(pointPar);
            }
            if (listPar.Count == 0)
                return null;
            else
                return listPar;
        }

        //获取段落的图片
        /// <summary>
        /// 获取段落的图片
        /// </summary>
        /// <param name="pointPar">段落对象</param>
        /// <param name="pictureIndex">图片在段落的索引</param>
        /// <returns>图片对象</returns>
        private Picture getParagraphPictures(Paragraph pointPar, int pictureIndex)
        {
            Picture pic = null;
            if (pointPar.Pictures.Count >= pictureIndex)
            {
                pic = pointPar.Pictures[pictureIndex];                
            }
            return pic;

        }

        

        /// <summary>
        /// 替换字符串（替换标记）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="newValue"></param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="isAllReplace">是否全部替换,true为全部替换,false为单个替换</param>
        /// <returns></returns>
        public Boolean ReplaceFlag(DocX doc, string replaceFlag, string newValue, string alignment, bool isAllReplace)
        {

            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);
            while (p != null)
            {
                try
                {
                    p.ReplaceText(replaceFlag, newValue);
                }
                catch (System.NullReferenceException e)
                {
                    System.Console.WriteLine(e.ToString());
                    continue;
                }

                if (isAllReplace == true)
                {
                    p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);
                }
                else
                {
                    p = null;
                    break;
                }
            }
            return true;
        }

        /// <summary>
        /// 替换字符串标记
        /// </summary>
        /// <param name="fromPath">打开文件路径</param>
        /// <param name="toPath">保存路径</param>
        /// <param name="replaceFlag">替换标记</param>
        /// <param name="newValue">新值</param>
        /// <param name="alignment">对齐方式,默认左对齐</param>
        /// <param name="isAllReplace">是否全部替换</param>
        /// <returns>是否成功</returns>
        public Boolean ReplaceFlag(string fromPath, string toPath, string replaceFlag, string newValue, string alignment, bool isAllReplace)
        {
            DocX doc = DocX.Load(fromPath);
            if (alignment.Equals(""))
            {
                alignment = null;
            }
            Boolean bl = ReplaceFlag(doc, replaceFlag, newValue, alignment, isAllReplace);
            doc.SaveAs(toPath);
            return bl;
        }

        //替换图片,使用原始尺寸
        /// <summary>
        /// 替换图片,使用原始尺寸
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="parIndex">有图片的段落的索引</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="alignment">对齐方式</param>
        /// <returns></returns>
        private Picture replacePicture(DocX doc, int parIndex, string imgPath, string alignment)
        {
            List<Paragraph> listPar = getPictureParagraphs(doc);
            if (listPar == null)
                return null;
            if (listPar.Count < parIndex)
                return null;
            if (getParagraphPictures(listPar[parIndex], 0) == null)
                return null;

            Novacode.Image img = null;
            try
            {
                img = doc.AddImage(imgPath);
            }
            catch (System.InvalidOperationException e)
            {
                classLims_NPOI.WriteLog(e, "");
                return null;
            }
            Picture pic = img.CreatePicture();

            Paragraph par = listPar[parIndex];
            par.Pictures[0].Remove();

            //par.Pictures.Add(pic);
            par.InsertPicture(pic);

            pic.Height = Convert.ToInt32(Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            pic.Width = Convert.ToInt32(Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));

            return pic;
        }

        //替换图片,使用指定的尺寸
        /// <summary>
        /// 替换图片,使用指定的尺寸
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="parIndex">有图片的段落的索引</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="height">图片高度</param>
        /// <param name="width">图片宽度</param>
        private void replacePicture(DocX doc, int parIndex, string imgPath, string alignment, double height, double width)
        {
            Picture pic = replacePicture(doc, parIndex, imgPath, alignment);

            if (pic == null)
            {
                return;
            }

            if (Convert.ToInt32(height) == 0 && Convert.ToInt32(width) == 0)
            {
                return;
            }
            else if (Convert.ToInt32(height) == 0)
            {
                height = Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * width;
            }
            else if (Convert.ToInt32(width) == 0)
            {
                width = Convert.ToDouble(pic.Width) / Convert.ToDouble(pic.Height) * height;
            }

            pic.Height = Convert.ToInt32(height);
            pic.Width = Convert.ToInt32(width);
        }

        //替换指定段落的第一个图片
        /// <summary>
        /// 替换指定段落的第一个图片, 图片高度或宽度为0时使用图片的原始尺寸
        /// </summary>
        /// <param name="fromPath">文件路径</param>
        /// <param name="parIndex">有图片的段落的索引</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="height">图片高度</param>
        /// <param name="width">图片宽度</param>
        public void replacePicture(string fromPath, int parIndex, string imgPath, string alignment, double height, double width)
        {
            if (File.Exists(fromPath))
            {
                Novacode.DocX doc = Novacode.DocX.Load(fromPath);
                if (height == 0 || width == 0)
                {
                    replacePicture(doc, parIndex, imgPath, alignment);
                }
                else
                {
                    replacePicture(doc, parIndex, imgPath, alignment, height, width);
                }
                doc.Save();
            }
            return;
        }

        //合并
        /// <summary>
        /// 合并多个word,按页
        /// </summary>
        /// <param name="tempDoc">word活动页</param>
        /// <param name="lastDoc">word尾页</param>
        /// <param name="outDoc">保存路径</param>
        /// <param name="pageCount">页数</param>
        public void getMergeTempByCount(string tempDoc, string lastDoc, string outDoc, int pageCount)
        {
            try
            {
                DocX oldDocument = DocX.Load(tempDoc);
                //DocX tempDocument = DocX.Load(tempDoc);
                DocX newDocument = DocX.Load(tempDoc);
                DocX lastDocument = DocX.Load(lastDoc);

                //reportno = fine + reportno;

                if (pageCount == 1)
                {
                    ////替换报告单号
                    //ReplaceFlag(lastDocument, reportFlag, reportno, null, true);
                    ////替换总页数
                    //ReplaceFlag(lastDocument, totalPageFlag, (startPage).ToString(), null, true);
                    ////替换当前页数
                    //ReplaceFlag(lastDocument, nowPageFlag, (startPage).ToString(), null, true);

                    lastDocument.SaveAs(outDoc);
                    return;
                }

                else
                {
                    int i = 1;
                    for (i = 1; i <= pageCount - 2; i++)
                    {
                        ////还原临时文件
                        //newDocument = tempDocument;
                        ////替换报告单号
                        //ReplaceFlag(newDocument, reportFlag, reportno, null, true);
                        ////替换总页数
                        //ReplaceFlag(newDocument, totalPageFlag, (pageCount + startPage - 1).ToString(), null, true);
                        ////替换当前页数
                        //ReplaceFlag(newDocument, nowPageFlag, (startPage + i).ToString(), null, true);
                        oldDocument.InsertDocument(newDocument);
                    }

                    ////写入首页
                    //ReplaceFlag(oldDocument, reportFlag, reportno, null, true);
                    //ReplaceFlag(oldDocument, totalPageFlag, (pageCount + startPage - 1).ToString(), null, true);
                    //ReplaceFlag(oldDocument, nowPageFlag, (startPage).ToString(), null, true);

                    ////写入尾页
                    //ReplaceFlag(lastDocument, reportFlag, reportno, null, true);
                    //ReplaceFlag(lastDocument, totalPageFlag, (pageCount + startPage - 1).ToString(), null, true);
                    //ReplaceFlag(lastDocument, nowPageFlag, (pageCount + startPage - 1).ToString(), null, true);

                    oldDocument.InsertDocument(lastDocument);
                    oldDocument.SaveAs(outDoc);
                    return;
                }
            }
            catch (Exception ex)
            {
                classLims_NPOI.WriteLog(ex, "");
                return;
            }
        }


        //插入图片（用户自定义图片尺寸）, LEFT,RIGHT,CENTER
        /// <summary>
        /// 插入图片（用户自定义图片尺寸）
        /// </summary>
        /// <param name="docPath">文件路径</param>
        /// <param name="replaceFlag">替换标记</param>
        /// <param name="imgPath">图片路径</param>
        /// <param name="alignment">水平对齐方式,可以是:LEFT,RIGHT,CENTER</param>
        /// <param name="height">图片高度,为0时使用原始尺寸</param>
        /// <param name="width">图片宽度,为0时使用原始尺寸</param>
        public void InsertPicture(string docPath, string replaceFlag, string imgPath, string alignment, double height, double width)
        {
            if (File.Exists(docPath))
            {
                Novacode.DocX doc = Novacode.DocX.Load(docPath);
                if(height == 0 || width == 0)
                {                    
                    InsertPicture(doc, replaceFlag, imgPath, alignment);
                }
                else
                {
                    InsertPicture(doc, replaceFlag, imgPath, alignment, height, width);
                }
                doc.Save();
            }
            return;
        }

        /// <summary>
        /// 插入图片（用户自定义图片尺寸）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="imgPath"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        public void InsertPicture(DocX doc, string replaceFlag, string imgPath, string alignment, double height, double width)
        {
            Picture pic = InsertPicture(doc, replaceFlag, imgPath, alignment);

            if (pic == null)
            {
                return;
            }

            if (Convert.ToInt32(height) == 0 && Convert.ToInt32(width) == 0)
            {
                return;
            }
            else if (Convert.ToInt32(height) == 0)
            {
                height = Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * width;
            }
            else if (Convert.ToInt32(width) == 0)
            {
                width = Convert.ToDouble(pic.Width) / Convert.ToDouble(pic.Height) * height;
            }

            pic.Height = Convert.ToInt32(height);
            pic.Width = Convert.ToInt32(width);

        }

        /// <summary>
        /// 插入图片（对图片尺寸没有要求）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="imgPath"></param>
        public Picture InsertPicture(DocX doc, string replaceFlag, string imgPath, string alignment)
        {
            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);

            if (p == null)
            {
                return null;
            }

            p.ReplaceText(replaceFlag, "");

            Novacode.Image img = null;
            try
            {
                img = doc.AddImage(imgPath);
            }
            catch (System.InvalidOperationException e)
            {
                classLims_NPOI.WriteLog(e, "");
                return null;
            }

            Picture pic = img.CreatePicture();

            //p.AppendPicture(pic);
            p.InsertPicture(pic);

            pic.Height = Convert.ToInt32(Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            pic.Width = Convert.ToInt32(Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            return pic;
        }

        #endregion

    }
}
