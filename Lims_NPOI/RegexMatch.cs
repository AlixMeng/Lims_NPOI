using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;

namespace nsLims_NPOI
{
    /// <summary>
    /// 正则匹配
    /// </summary>
    class RegexMatch
    {
        public RegexMatch()
        {
        }

        //取出上标
        public static object[] RegexUp(string cellValue)
        {
            Regex reg = new Regex(@"\^\(.+?\)");
            MatchCollection mc = reg.Matches(cellValue);
            List<Point> lp = new List<Point>();
            while (mc.Count > 0)//一直获取,直到全部取出
            {
                Match mtch = mc[0];
                string upVal = mtch.Value;
                int position = mtch.Index;                
                int length = mtch.Length-3;//去掉标记后长度应该-3
                lp.Add(new Point(position,length));
                cellValue = Replace(cellValue, upVal, upVal.Substring(2, upVal.Length - 3));
                mc = reg.Matches(cellValue);
            }
            return new object[] { cellValue, lp, "UP" };
        }

        //取出下标
        public static object[] RegexDown(string cellValue)
        {
            Regex reg = new Regex(@"_\(.+?\)");
            MatchCollection mc = reg.Matches(cellValue);
            List<Point> lp = new List<Point>();
            while (mc.Count > 0)//一直获取,直到全部取出
            {
                Match mtch = mc[0];
                string downVal = mtch.Value;
                int position = mtch.Index;
                int length = mtch.Length-3;//去掉标记后长度应该-3
                lp.Add(new Point(position, length));
                cellValue = Replace(cellValue, downVal, downVal.Substring(2, downVal.Length - 3));
                mc = reg.Matches(cellValue);
            }
            return new object[] { cellValue, lp, "DOWN" };
        }

        //取出上下标
        public static object[] RegexUpAndDown(string cellValue)
        {
            #region 先查找上标
            Regex reg = new Regex(@"\^\(.+?\)");
            MatchCollection mc = reg.Matches(cellValue);
            List<object[]> lp = new List<object[]>();
            for (int i=0;i< mc.Count; i++)//一直获取,直到全部取出
            {
                Match mtch = mc[i];
                int position = mtch.Index;
                int length = mtch.Length;
                lp.Add(new object[] { position, length, "UP" });
              
            }
            #endregion 先查找上标

            #region 再查找下标
            reg = new Regex(@"_\(.+?\)");
            mc = reg.Matches(cellValue);
            for (int i = 0; i < mc.Count; i++)//一直获取,直到全部取出
            {
                Match mtch = mc[i];
                int position = mtch.Index;
                int length = mtch.Length;
                lp.Add(new object[] { position, length, "DOWN" });

            }
            #endregion 再查找下标

            #region 匹配数组的位置(第一列的position)排序
            List<object[]> lpNew = new List<object[]>();
            lpNew = lp.OfType<object[]>().OrderBy(e => e[0]).ToList();//根据第一列排序
            #endregion 匹配数组的位置排序

            //去掉标记字符并修正位置和长度
            int np = 0;//累计位置
            for(int i=0;i<lpNew.Count;i++, np++)
            {
                //剔除标记字符*_()
                string matchStr = cellValue.Substring((int)lpNew[i][0] - (np * 3), (int)lpNew[i][1]);
                string replaceStr = matchStr.Substring(2, matchStr.Length - 3);
                cellValue = Replace(cellValue, matchStr, replaceStr);
                lpNew[i][0] = (int)lpNew[i][0] - (np*3);
                lpNew[i][1] = (int)lpNew[i][1] - 3;
            }
            return new object[] { cellValue, lpNew};
        }

        /// <summary>
        /// 只替换第一个匹配字符
        /// </summary>
        /// <param name="source"></param>
        /// <param name="match"></param>
        /// <param name="replacement"></param>
        /// <returns></returns>
        public static string Replace(string source, string match, string replacement)
        {
            char[] sArr = source.ToCharArray();
            char[] mArr = match.ToCharArray();
            char[] rArr = replacement.ToCharArray();
            int idx = IndexOf(sArr, mArr);
            if (idx == -1)
            {
                return source;
            }
            else
            {
                return new string(sArr.Take(idx).Concat(rArr).Concat(sArr.Skip(idx + mArr.Length)).ToArray());
            }
        }

        /// <summary>
        /// 查找字符数组在另一个字符数组中匹配的位置
        /// </summary>
        /// <param name="source">源字符数组</param>
        /// <param name="match">匹配字符数组</param>
        /// <returns>匹配的位置，未找到匹配则返回-1</returns>
        private static int IndexOf(char[] source, char[] match)
        {
            int idx = -1;
            for (int i = 0; i < source.Length - match.Length; i++)
            {
                if (source[i] == match[0])
                {
                    bool isMatch = true;
                    for (int j = 0; j < match.Length; j++)
                    {
                        if (source[i + j] != match[j])
                        {
                            isMatch = false;
                            break;
                        }
                    }
                    if (isMatch)
                    {
                        idx = i;
                        break;
                    }
                }
            }
            return idx;
        }


    }    
}
