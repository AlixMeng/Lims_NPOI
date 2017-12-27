using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;
using System.IO;
using System.Data;
using WORD = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Xml;
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
        /// 加载DocX
        /// </summary>
        /// <returns></returns>
        public DocX Load(string fileName)
        {
            return DocX.Load(fileName);
        }

        /// <summary>
        /// 保存DocX
        /// </summary>
        /// <param name="doc"></param>
        public void Save(DocX doc)
        {
            doc.Save();
        }


        /// <summary>
        /// 生成一个表格
        /// </summary>
        /// <param name="fromPath">源文件</param>
        /// <param name="strXml">xml格式字符串</param>
        /// <param name="replaceFlag">替换标记,定位表格位置</param>
        /// <param name="bEmptyParagraph"></param>
        /// <param name="sPercent">各列所占表格的百分比,逗号隔开</param>
        /// <param name="sPagePercent">表格所占页面的百分比</param>
        /// <returns></returns>
        public Table GenerateTable(string fromPath, string strXml, string replaceFlag,
            bool bEmptyParagraph, string sPercent, string sPagePercent)
        {
            DocX doc = DocX.Load(fromPath);
            DataTable dt = XMLDeserialize(strXml);
            int nRow = dt.Rows.Count;
            int nCol = dt.Columns.Count;
            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, null);
            if (p == null)
            {
                return null;
            }
            Table tbl = p.InsertTableBeforeSelf(nRow, nCol);

            int[] aRowNCol = new int[nRow];//每一行有多少列

            #region 数组：每行有多少列
            //目的是由于合并情况的出现会导致每行列数减少，但是合并单元格以后行不会减少
            for (int i = 0; i < aRowNCol.Length; i++)
            {
                aRowNCol[i] = nCol;
            }
            #endregion 数组：每行有多少列

            #region 合并单元格,从右下角往左上角扫
            int nLe;
            int nUp;

            for (int j = nCol - 1; j >= 0; j--)
            {
                for (int i = nRow - 1; i >= 0; i--)
                {
                    nUp = 0;
                    nLe = 0;
                    if (dt.Rows[i][j].ToString().ToUpper().Trim() == "UP" || dt.Rows[i][j].ToString().ToUpper().Trim() == "LE")
                    {
                        continue;
                    }
                    else
                    {
                        tbl.Rows[i].Cells[j].Paragraphs[0].Append(dt.Rows[i][j].ToString());
                    }

                    for (int n = i + 1; n < nRow; n++)
                    {
                        if (dt.Rows[n][j].ToString().ToUpper().Trim() == "UP")
                        {
                            aRowNCol[n] -= 1;
                            nUp++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    for (int m = j + 1; m < aRowNCol[i]; m++)
                    {
                        if (dt.Rows[i][m].ToString().ToUpper().Trim() == "LE")
                        {
                            nLe++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    #region 合并行
                    if (nUp > 0)
                    {
                        try
                        {
                            tbl.MergeCellsInColumn(j, i, i + nUp);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            classLims_NPOI.WriteLog(e, "");
                            continue;
                        }
                    }
                    #endregion 合并行
                    if (nLe > 0)
                    {
                        aRowNCol[i] -= nLe;
                        try
                        {
                            tbl.Rows[i].MergeCells(j, j + nLe);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            classLims_NPOI.WriteLog(e, "");
                            continue;
                        }

                        #region 删除合并单元格中的多余回车
                        for (int z = 0; z < nLe; z++)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[tbl.Rows[i].Cells[j].Paragraphs.Count - 1].Remove(false);
                        }
                        #endregion 删除合并单元格中的多余回车
                        if (nUp > 0)
                        {
                            for (int l = i + 1; l <= i + nUp; l++)
                            {
                                try
                                {
                                    tbl.Rows[l].MergeCells(j, j + nLe);
                                }
                                catch (System.ArgumentOutOfRangeException e)
                                {
                                    classLims_NPOI.WriteLog(e, "");
                                    continue;
                                }
                            }
                        }
                    }
                    #region 科学计数法 & 上下标
                    int nParagraphs = tbl.Rows[i].Cells[j].Paragraphs.Count;//看这个单元格有多少paragraphs
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        String sComment = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper();

                        if (sComment.IndexOf("E") > 0)
                        {
                            string sPre = sComment.Substring(0, sComment.IndexOf("E"));
                            string sZhishu = sComment.Substring(sComment.IndexOf("E") + 1, sComment.Length - sComment.IndexOf("E") - 1);
                            if ((Regex.IsMatch(sPre, @"^\d+\.\d+$") || Regex.IsMatch(sPre, @"^\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+\.\d+$"))
                                    && (Regex.IsMatch(sZhishu, @"^\d+$") || Regex.IsMatch(sZhishu, @"^[-]+\d+$")))
                            {
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(0, sComment.Length);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(Convert.ToDecimal(sPre).ToString() + "×10{U|" + Int32.Parse(sZhishu).ToString() + "}");
                            }
                        }
                    }
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        #region 上标
                        int iBeginUp = -1;
                        Novacode.Formatting formattingUp = new Novacode.Formatting();
                        formattingUp.Script = Script.superscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|") > 0)
                        {
                            iBeginUp = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|");
                            int iEndUp = -1;
                            for (iEndUp = iBeginUp + 3; iEndUp < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndUp++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndUp, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginUp + 3) != iEndUp)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginUp + 3, iEndUp - iBeginUp - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginUp, iEndUp - iBeginUp + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginUp, strSub, false, formattingUp);
                            }
                        }
                        #endregion 上标

                        #region 下标
                        int iBeginDown = -1;
                        Novacode.Formatting formattingDown = new Novacode.Formatting();
                        formattingDown.Script = Script.subscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|") > 0)
                        {
                            iBeginDown = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|");
                            int iEndDown = -1;
                            for (iEndDown = iBeginDown + 3; iEndDown < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndDown++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndDown, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginDown + 3) != iEndDown)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginDown + 3, iEndDown - iBeginDown - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginDown, iEndDown - iBeginDown + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginDown, strSub, false, formattingDown);
                            }
                        }
                        #endregion 下标
                    }
                    #endregion 上下标
                }
            }


            #endregion 合并单元格,从右下角往左上角扫

            if (bEmptyParagraph)
            {
                Paragraph p1 = tbl.InsertParagraphAfterSelf("");
                //p1.Remove();

            }
            Paragraph rp = GetParagraphByReplaceFlag(doc, replaceFlag, "LEFT");
            rp.ReplaceText(replaceFlag, "");
            SetTableBorderLine(tbl, "BOTTOM");
            SetTableBorderLine(tbl, "TOP");
            SetTableBorderLine(tbl, "RIGHT");
            SetTableBorderLine(tbl, "LEFT");
            SetTableBorderLine(tbl, "INSIDEV");
            SetTableBorderLine(tbl, "INSIDEH");
            SetTableColWidth(fromPath, tbl, sPercent, sPagePercent);
            doc.Save();
            return tbl;
        }

        /// <summary>
        /// 设置表格列宽（页面比例）
        /// </summary>
        /// <param name="fromPath"></param>
        /// <param name="tbl"></param>
        /// <param name="sPercent">各列所占表格的百分比</param>
        /// <param name="sPagePercent">表格所占页面的百分比</param>
        /// <returns></returns>
        public Table SetTableColWidth(string fromPath, Table tbl, string sPercent, string sPagePercent)
        {
            DocX doc = DocX.Load(fromPath);
            string[] aPercent = sPercent.Split(',');
            if (tbl.ColumnCount != aPercent.Length)
            {
                return tbl;
            }
            double nPagePercent = Convert.ToDouble(sPagePercent) / 100;
            double sum_aPercent = 0.0;
            for (int i = 0; i < aPercent.Length; i++)
            {
                sum_aPercent += Convert.ToDouble(aPercent[i]);
            }
            float[] colWidths = new float[aPercent.Length];
            for (int i = 0; i < aPercent.Length; i++)
            {
                colWidths[i] = (float)(Convert.ToDouble(aPercent[i]) / sum_aPercent * nPagePercent * doc.PageWidth);
            }
            tbl.SetWidths(colWidths);

            tbl.Alignment = Alignment.center;
            tbl.AutoFit = AutoFit.ColumnWidth;

            return tbl;
        }

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
        /// <param name="alignment">对齐方式</param>
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

        /// <summary>
        /// 设置表格边框
        /// </summary>
        /// <param name="tbl"></param>
        /// <param name="sDirections">要加的边框位置,(LEFT,TOP,BOTTOM,RIGHT)</param>
        /// <returns></returns>
        public Table SetTableBorderLine(Table tbl, string sDirections)
        {
            if (tbl == null)
            {
                return null;
            }

            string[] aDirections = sDirections.Split(',');
            for (int i = 0; i < aDirections.Length; i++)
            {
                if (aDirections[i].ToUpper() == "LEFT")
                {
                    for (int j = 0; j < tbl.RowCount; j++)
                    {
                        SetTableCellStyle(tbl, j, 0, "BORDER:LEFT;");
                    }
                }

                if (aDirections[i].ToUpper() == "TOP")
                {
                    for (int j = 0; j < tbl.Rows[0].Cells.Count; j++)
                    {
                        SetTableCellStyle(tbl, 0, j, "BORDER:TOP;");
                    }
                }

                if (aDirections[i].ToUpper() == "BOTTOM")
                {
                    for (int j = 0; j < tbl.Rows[tbl.RowCount - 1].Cells.Count; j++)
                    {
                        SetTableCellStyle(tbl, tbl.RowCount - 1, j, "BORDER:BOTTOM;");
                    }
                }

                if (aDirections[i].ToUpper() == "RIGHT")
                {
                    for (int j = 0; j < tbl.RowCount; j++)
                    {
                        SetTableCellStyle(tbl, j, tbl.Rows[j].Cells.Count - 1, "BORDER:RIGHT;");
                    }
                }

                if (aDirections[i].ToUpper() == "INSIDEH")
                {
                    for (int j = 0; j < tbl.RowCount - 1; j++)
                    {
                        for (int k = 0; k < tbl.Rows[j].Cells.Count; k++)
                        {
                            SetTableCellStyle(tbl, j, k, "BORDER:BOTTOM;");
                        }
                    }
                }

                if (aDirections[i].ToUpper() == "INSIDEV")
                {
                    for (int j = 0; j < tbl.RowCount; j++)
                    {
                        for (int k = 0; k < tbl.Rows[j].Cells.Count - 1; k++)
                        {
                            SetTableCellStyle(tbl, j, k, "BORDER:RIGHT;");
                        }
                    }
                }
            }
            return tbl;
        }

        public Table SetTableCellStyle(Table tbl, int x, int y, string sCellStyleAll)
        {
            if (x >= tbl.Rows.Count || x < 0)
            {
                return tbl;
            }
            else
            {
                if (y >= tbl.Rows[x].Cells.Count || (y != -1 && y < 0))
                {
                    return tbl;
                }
            }
            string[] aCellStyle;//由属性串分割成的属性数组
            string sStyleName = "";
            string sStyleValue = "";

            aCellStyle = sCellStyleAll.Split(';');
            for (int m = 0; m < aCellStyle.Length; m++)
            {
                if (aCellStyle[m].IndexOf(':') < 0)
                {
                    continue;
                }

                //sStyleName = aCellStyle[m].Substring(0, aCellStyle[m].IndexOf(':'));
                //sStyleValue = aCellStyle[m].Substring(aCellStyle[m].IndexOf(':') + 1);
                string[] aCellStyleOne = aCellStyle[m].Split(':');
                sStyleName = aCellStyleOne[0];
                sStyleValue = aCellStyleOne[1];

                BorderStyle BorderStyle_Tcbs = BorderStyle.Tcbs_single;
                //BorderStyle BorderStyle_Tcbs = BorderStyle.Tcbs_dotted;
                BorderSize BorderStyle_Size = BorderSize.one;
                int BorderStyle_Space = 0;
                System.Drawing.Color BorderStyle_Color = System.Drawing.Color.Black;
                TableCellBorderType tblCellBorderType = TableCellBorderType.Left;

                if (aCellStyleOne.Length > 2)
                {
                    for (int p = 2; p < aCellStyleOne.Length; p++)
                    {
                        string sBorderProperty = aCellStyleOne[p].Substring(0, aCellStyleOne[p].IndexOf('_'));
                        string sBorderValue = aCellStyleOne[p].Substring(aCellStyleOne[p].IndexOf('_') + 1);
                        if (sBorderProperty.ToUpper() == "BORDERSTYLE")//调边线样式
                        {
                            if (sBorderValue.ToUpper() == "DOTTED")
                            {
                                BorderStyle_Tcbs = BorderStyle.Tcbs_dotted;
                            }
                            else if (sBorderValue.ToUpper() == "NONE")
                            {
                                BorderStyle_Tcbs = BorderStyle.Tcbs_none;
                            }
                            else if (sBorderValue.ToUpper() == "SINGLE")
                            {
                                BorderStyle_Tcbs = BorderStyle.Tcbs_single;
                            }
                        }
                        else if (sBorderProperty.ToUpper() == "BORDERSIZE")//调边线粗细
                        {

                        }
                        else if (sBorderProperty.ToUpper() == "BORDERCOLOR")//调边线颜色
                        {
                            if (sBorderValue.ToUpper() == "RED")
                            {
                                BorderStyle_Color = System.Drawing.Color.Red;
                            }
                            else if (sBorderValue.ToUpper() == "WHITE")
                            {
                                BorderStyle_Color = System.Drawing.Color.White;
                            }
                        }
                    }
                }

                if (sStyleName.ToUpper() == "BORDER")
                {
                    if (sStyleValue.ToUpper() == "LEFT")
                    {
                        tblCellBorderType = TableCellBorderType.Left;
                    }
                    else if (sStyleValue.ToUpper() == "TOP")
                    {
                        tblCellBorderType = TableCellBorderType.Top;
                    }
                    else if (sStyleValue.ToUpper() == "RIGHT")
                    {
                        tblCellBorderType = TableCellBorderType.Right;
                    }
                    else if (sStyleValue.ToUpper() == "BOTTOM")
                    {
                        tblCellBorderType = TableCellBorderType.Bottom;
                    }

                    if (y == -1)
                    {
                        for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                        {
                            tbl.Rows[x].Cells[z].SetBorder(tblCellBorderType, new Border(BorderStyle_Tcbs, BorderStyle_Size, BorderStyle_Space, BorderStyle_Color));
                        }
                    }
                    else
                    {
                        tbl.Rows[x].Cells[y].SetBorder(tblCellBorderType, new Border(BorderStyle_Tcbs, BorderStyle_Size, BorderStyle_Space, BorderStyle_Color));
                    }
                }

                if (sStyleName.ToUpper() == "PARAGRAPHALIGN")
                {
                    if (sStyleValue.ToUpper() == "LEFT")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.left;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.left;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "RIGHT")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.right;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.right;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "CENTER")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.center;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.center;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "VTOP")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].VerticalAlignment = VerticalAlignment.Top;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].VerticalAlignment = VerticalAlignment.Top;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "VCENTER")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].VerticalAlignment = VerticalAlignment.Center;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].VerticalAlignment = VerticalAlignment.Center;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "VBOTTOM")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].VerticalAlignment = VerticalAlignment.Bottom;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].VerticalAlignment = VerticalAlignment.Bottom;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "BOTH")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.both;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.both;
                        }
                    }
                }
            }

            return tbl;
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
        /// <param name="alignment">对齐方式</param>
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
        /// <param name="alignment">对齐方式</param>
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

            Novacode.Picture pic = img.CreatePicture();

            //p.AppendPicture(pic);
            p.InsertPicture(pic);

            pic.Height = Convert.ToInt32(Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            pic.Width = Convert.ToInt32(Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            return pic;
        }

        //added by LIUJIE 2017-09-18
        /// <summary>
        /// 插入图片及图片注释（对图片尺寸没有要求）
        /// </summary>
        /// <param name="oldPath">添加的doc路径</param>
        /// <param name="oPath">添加图片的数组</param>
        /// <param name="replaceFlag">替换符</param>
        /// <param name="oRemark">图片备注数组</param>
        public void AddWordPic(string oldPath, object[] oPath, string replaceFlag, object[] oRemark)
        {
            DocX oldDocument = DocX.Load(oldPath);
            Paragraph ss = null;
            Novacode.Image img = null;
            ss = GetParagraphByReplaceFlag(oldDocument, replaceFlag, "CENTER");
            ss.ReplaceText(replaceFlag, "");
            if (!(oPath == null || oRemark == null))
            {
                try
                {
                    string[] imagePath = classLims_NPOI.dArray2String1(oPath);
                    string[] remark = classLims_NPOI.dArray2String1(oRemark);
                    for (int i = 0; i < imagePath.Length; i++)
                    {
                        img = oldDocument.AddImage(imagePath[i]);
                        Picture pic = img.CreatePicture();
                        ss.AppendPicture(pic);
                        pic.Height = Convert.ToInt32(Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * Convert.ToDouble(oldDocument.PageWidth));
                        pic.Width = Convert.ToInt32(Convert.ToDouble(oldDocument.PageWidth));
                        ss.AppendLine(remark[i] + "\n");
                        //ss.AppendLine("\n");
                        ss.Alignment = Alignment.center;
                    }
                }
                catch (System.InvalidOperationException e)
                {
                    classLims_NPOI.WriteLog(e, "");
                    return;
                }
            }
            oldDocument.Save();
            return;
        }

        /// <summary>
        /// Xml反序列化为DataTable
        /// </summary>
        /// <param name="strXml"></param>
        /// <returns></returns>
        public DataTable XMLDeserialize(string strXml)
        {
            classLims_NPOI.WriteLog(strXml, "");
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(strXml);
            XmlNode xn = xmlDoc.SelectSingleNode("complexType");
            XmlNodeList xnl = xn.ChildNodes;
            string sCol = ((XmlElement)xnl[0]).GetAttribute("length");
            int nCol = int.Parse(sCol);
            DataTable dt = new DataTable();

            for (int i = 0; i < nCol; i++)
            {
                dt.Columns.Add();
            }
            int tmpCol = 0;
            foreach (XmlNode xnf in xnl)
            {
                DataRow dr = dt.NewRow();
                tmpCol = 0;

                XmlElement xe = (XmlElement)xnf;

                XmlNodeList xnf1 = xe.ChildNodes;
                foreach (XmlNode xn2 in xnf1)
                {
                    dr[tmpCol++] = xn2.InnerText;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }


        #endregion

    }
}
