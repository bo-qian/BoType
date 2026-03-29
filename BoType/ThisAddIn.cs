using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace BoType
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// 插入行内公式
        /// </summary>
        public void InsertInlineEquation()
        {
            Word.Application app = this.Application;
            if (app.Documents.Count == 0) return;

            Word.Document doc = app.ActiveDocument;
            Word.Selection selection = app.Selection;

            Word.Range rng = selection.Range;
            rng.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            doc.OMaths.Add(rng);
            Word.OMath mathObj = rng.OMaths[1];

            // 设为行内公式
            mathObj.Type = Word.WdOMathType.wdOMathInline;

            Word.Range inputRng = mathObj.Range;
            inputRng.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            inputRng.Select();
        }

        private bool TryGetChapterInfo(Word.Document doc, Word.Selection selection, out bool hasAutoNum, out int chapterCount)
        {
            hasAutoNum = false;
            chapterCount = 0;

            Word.Range findRng = selection.Range.Duplicate;
            findRng.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            findRng.Find.ClearFormatting();
            findRng.Find.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            findRng.Find.Text = "";
            findRng.Find.Forward = false;
            findRng.Find.Wrap = Word.WdFindWrap.wdFindStop;
            findRng.Find.Format = true;

            if (!findRng.Find.Execute())
            {
                return false;
            }

            string listStr = null;
            try { listStr = findRng.ListFormat.ListString; } catch { }

            if (!string.IsNullOrWhiteSpace(listStr) && listStr.Any(char.IsDigit))
            {
                hasAutoNum = true;
            }
            else
            {
                hasAutoNum = false;

                string headingText = null;
                try { headingText = findRng.Text?.Trim(); } catch { }
                bool hasExplicitNum = false;

                if (!string.IsNullOrEmpty(headingText))
                {
                    var match = System.Text.RegularExpressions.Regex.Match(headingText, @"^(?:第)?\s*(\d+)");
                    if (match.Success && int.TryParse(match.Groups[1].Value, out int explicitNum))
                    {
                        chapterCount = explicitNum;
                        hasExplicitNum = true;
                    }
                }

                if (!hasExplicitNum)
                {
                    Word.Range countRng = doc.Range(0, findRng.End);
                    countRng.Find.ClearFormatting();
                    countRng.Find.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                    countRng.Find.Text = "";
                    countRng.Find.Forward = true;
                    countRng.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    countRng.Find.Format = true;

                    int count = 0;
                    while (countRng.Find.Execute())
                    {
                        if (countRng.Start > findRng.Start) break;
                        count++;
                    }
                    chapterCount = count > 0 ? count : 1;
                }
            }
            return true;
        }

        /// <summary>
        /// 插入单行编号公式 (三栏表格模式)
        /// </summary>
        /// <param name="numberStyle">0: 无, 1: (1), 2: (1.1), 3: (1-1)</param>
        public void InsertNumberedEquation(int numberStyle = 1, float sideWidth = 38.0f)
        {
            Word.Application app = this.Application;
            if (app.Documents.Count == 0) return;

            Word.Document doc = app.ActiveDocument;
            Word.Selection selection = app.Selection;

            bool hasAutoNum = true;
            int chapterCount = 1;
            if (numberStyle > 1)
            {
                if (!TryGetChapterInfo(doc, selection, out hasAutoNum, out chapterCount))
                {
                    System.Windows.Forms.MessageBox.Show("当前位置前面没有任何一级标题，不允许使用带章节的编号。", "BoType - 错误");
                    return;
                }
            }

            if (numberStyle == 0)
            {
                // 选择“无”时直接插入原生单行公式，不生成表格
                Word.Range rng = selection.Range;
                rng.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                doc.OMaths.Add(rng);

                Word.OMath mathObjNative = rng.OMaths[1];
                mathObjNative.Type = Word.WdOMathType.wdOMathDisplay;

                Word.Range inputRngNative = mathObjNative.Range;
                inputRngNative.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                inputRngNative.Select();
                return;
            }

            // 1. 动态获取当前光标所在节的页面设置
            Word.PageSetup pageSetup = selection.PageSetup;
            float pageWidth = pageSetup.PageWidth;
            float leftMargin = pageSetup.LeftMargin;
            float rightMargin = pageSetup.RightMargin;

            // 2. 计算可用的排版宽度
            float usableWidth = pageWidth - leftMargin - rightMargin;
            // 定义左右两侧专用于编号和占位的列宽
            float centerWidth = usableWidth - (sideWidth * 2);

            // 3. 在当前位置插入一个 1行3列 的无边框表格
            Word.Table table = doc.Tables.Add(selection.Range, 1, 3);
            // 去除表格边框
            table.Borders.Enable = 0;
            // 设置整行内容垂直居中
            table.Rows[1].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            // 禁用自动调整并重置表格整体缩进，防止其他文档样式干扰
            table.AllowAutoFit = false;
            table.Rows.LeftIndent = 0f;

            // 严格设定三等分/对称的两侧宽度，才能保证中间列绝对在页面中心
            table.Cell(1, 1).Width = sideWidth;
            table.Cell(1, 2).Width = centerWidth;
            table.Cell(1, 3).Width = sideWidth;

            // ==== 处理第一列：左侧占位 ====
            Word.Cell cellLeft = table.Cell(1, 1);
            // 取消左侧边距，让未来可能的填充物完全贴合页面最左侧
            cellLeft.LeftPadding = 0f;

            // ==== 处理第二列：中间区域插入 Display 模式公式 ====
            Word.Cell cellCenter = table.Cell(1, 2);
            // 中间列文本水平居中
            cellCenter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // 获取中间单元格的开始位置并选中（模拟用户在此处插入）
            Word.Range rngCenter = cellCenter.Range;
            rngCenter.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            rngCenter.Select();

            // 借助 Selection 直接无痕插入空白公式框。
            // 针对完全 Empty 的 Range，直接 OMaths.Add(Range) 容易导致返回的 Range 是空或 OMaths 集合长度为 0 的异常
            doc.OMaths.Add(app.Selection.Range);

            // 重新在单元格的作用域内获取被创建的公式
            Word.OMath mathObj = cellCenter.Range.OMaths[1];

            // 强制设为显示模式 (wdOMathDisplay = 0)，防止在此框架下变成压缩版的行内公式
            mathObj.Type = Word.WdOMathType.wdOMathDisplay;
            // mathObj.BuildUp() 被省略，因为对于空的公式框进行 BuildUp 生成二维结构没有必要且容易出错

            // ==== 处理第三列：右侧区域插入编号和括号 ====
            Word.Cell cellRight = table.Cell(1, 3);
            // 右边列文本右对齐
            cellRight.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            if (numberStyle > 0)
            {
                Word.Range rngRight = cellRight.Range;
                rngRight.Collapse(Word.WdCollapseDirection.wdCollapseStart);  // 定位到单元格开始

                // 插入括号
                rngRight.Text = "()";

                // 将光标定位在左右括号之间，准备插入编号的域代码
                Word.Range rngNum = table.Cell(1, 3).Range;
                rngNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                rngNum.Move(Word.WdUnits.wdCharacter, 1);  // 向右移动1个字符，进入 '(' 的右侧

                // 记录最开始的插入位置，供后续外围包裹书签使用
                int bmStartPos = rngNum.Start;

                Word.Range bmRng;

                if (numberStyle == 1)
                {
                    // 插入纯数字自动编号 (SEQ 域)
                    Word.Field field = doc.Fields.Add(rngNum, Word.WdFieldType.wdFieldEmpty, @"SEQ 公式 \* ARABIC", false);
                    Word.Range lastRng = field.Result.Duplicate;
                    lastRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    lastRng.Move(Word.WdUnits.wdCharacter, 1);
                    bmRng = doc.Range(bmStartPos, lastRng.Start);
                }
                else
                {
                    // 插入章节编号 (如 1.1)：组合 STYLEREF 和 SEQ 或者静态数字

                    Word.Range sepRng;
                    if (hasAutoNum)
                    {
                        // 1. 插入获取标题1编号的域 { STYLEREF "标题 1" \s }
                        Word.Field styleField = doc.Fields.Add(rngNum, Word.WdFieldType.wdFieldEmpty, @"STYLEREF ""标题 1"" \s", false);

                        // 定位到刚插入的 STYLEREF 域末尾，必须向右移动 1 个字符跳出域的右边界 '}'
                        sepRng = styleField.Result.Duplicate;
                        sepRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        sepRng.Move(Word.WdUnits.wdCharacter, 1); // 【关键修复】：跳出域的作用判定区
                    }
                    else
                    {
                        // 标题没有自动编号，直接使用静态的章节数字
                        rngNum.Text = chapterCount.ToString();
                        sepRng = rngNum.Duplicate;
                        sepRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }

                    // 2. 插入分隔符（这里根据选择不同赋予 . 或是 -）
                    sepRng.Text = (numberStyle == 2) ? "." : "-";

                    // 定位到分隔符后
                    Word.Range seqRng = sepRng.Duplicate;
                    seqRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // 3. 插入按一级标题重新计数的 SEQ 域 { SEQ 公式 \* ARABIC \s 1 }
                    // 注意 \s 1 意思是遇到"标题 1"级别样式就重新开始计数
                    Word.Field seqField = doc.Fields.Add(seqRng, Word.WdFieldType.wdFieldEmpty, @"SEQ 公式 \* ARABIC \s 1", false);

                    // 为了防止多域范围交叉引起书签在 F9 更新时因跨域而崩溃，需要将整个复合编号从域外的最外围进行完整包裹
                    Word.Range lastRng = seqField.Result.Duplicate;
                    lastRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    lastRng.Move(Word.WdUnits.wdCharacter, 1); // 同样跳出最后一个域的 '}'

                    bmRng = doc.Range(bmStartPos, lastRng.Start);
                }

                // 在插入出来的编号上加盖书签以支持交叉引用
                // 书签命名规则必须以 OLE_LINK 开头，Word 插入交叉引用后才能支持 Ctrl+左键 跳转
                string bookmarkName = "OLE_LINK" + Guid.NewGuid().ToString("N").Substring(0, 8); 
                doc.Bookmarks.Add(bookmarkName, bmRng);
            }

            // 【核心调整】：彻底消除表格右侧内边距，让编号完全贴合最右边距边缘
            cellRight.RightPadding = 0f;
            cellRight.Range.ParagraphFormat.RightIndent = 0f;
            cellRight.Range.ParagraphFormat.CharacterUnitRightIndent = 0f;

            // 【修复】：确保在写入纯文本和域之后，再对整个右侧单元格强制应用一次右对齐
            cellRight.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 顺便确保整个内容的垂直居中
            cellRight.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            // 确保编号不为斜体
            cellRight.Range.Font.Italic = 0;

            // 消除表格后下一个回车可能的斜体属性
            Word.Range afterTable = table.Range.Duplicate;
            afterTable.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterTable.Font.Italic = 0;

            // 4. 善后：不再增加多余的回车换行，将光标定位在中间列的空公式框内，方便用户直接输入
            Word.Range inputRng = mathObj.Range;
            inputRng.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            inputRng.Select();
        }

        public void SaveDefaultSettings(int styleIndex, float sideWidth)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\BoTypeAddIn"))
                {
                    key.SetValue("DefaultNumberStyle", styleIndex);
                    // 注册表默认存整型或字符串，这里可存为字符串
                    key.SetValue("DefaultSideWidth", sideWidth.ToString());
                }
            }
            catch { }
        }

        public void SaveDefaultNumberStyle(int styleIndex)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\BoTypeAddIn"))
                {
                    key.SetValue("DefaultNumberStyle", styleIndex);
                }
            }
            catch { }
        }

        public int LoadDefaultNumberStyle()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\BoTypeAddIn"))
                {
                    if (key != null)
                    {
                        return (int)(key.GetValue("DefaultNumberStyle", 1));
                    }
                }
            }
            catch { }
            return 1; // 默认为 (1)
        }

        public float LoadDefaultSideWidth()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\BoTypeAddIn"))
                {
                    if (key != null)
                    {
                        string val = key.GetValue("DefaultSideWidth") as string;
                        if (!string.IsNullOrEmpty(val) && float.TryParse(val, out float width))
                        {
                            return width;
                        }
                    }
                }
            }
            catch { }
            return 38.0f; // 默认为 38.0
        }

        public void WrapSelectedEquation(int numberStyle, float sideWidth = 38.0f)
        {
            if (numberStyle <= 0) return;

            Word.Application app = this.Application;
            if (app.Documents.Count == 0) return;

            Word.Document doc = app.ActiveDocument;
            Word.Selection selection = app.Selection;

            bool hasAutoNum = true;
            int chapterCount = 1;
            if (numberStyle > 1)
            {
                if (!TryGetChapterInfo(doc, selection, out hasAutoNum, out chapterCount))
                {
                    System.Windows.Forms.MessageBox.Show("当前位置前面没有任何一级标题，不允许使用带章节的编号。", "BoType - 错误");
                    return;
                }
            }

            Word.OMaths omaths = selection.OMaths;
            if (omaths.Count == 0 && selection.Paragraphs.Count > 0)
            {
                omaths = selection.Paragraphs[1].Range.OMaths;
            }
            if (omaths.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("当前光标未选中公式，请先点击需要编号的单行公式。", "BoType - 提示");
                return;
            }

            Word.OMath mathToWrap = omaths[1];

            // 检查是否已经在三栏表格中
            bool isAlreadyNumbered = false;
            Word.Table existingTable = null;
            if (mathToWrap.Range.Information[Word.WdInformation.wdWithInTable])
            {
                Word.Cell cell = null;
                try { cell = mathToWrap.Range.Cells[1]; } catch { }
                if (cell != null && cell.ColumnIndex == 2)
                {
                    existingTable = cell.Range.Tables[1];
                    if (existingTable.Columns.Count == 3)
                    {
                        isAlreadyNumbered = true;
                    }
                }
            }

            if (isAlreadyNumbered)
            {
                // 已有编号，修改编号样式并保留书签
                Word.Cell cellRightExisting = existingTable.Cell(1, 3);

                Word.Range rightRange = cellRightExisting.Range;
                rightRange.End -= 1; // 排除单元格结束符

                bool oldShowHidden = false;
                try { oldShowHidden = doc.Bookmarks.ShowHidden; doc.Bookmarks.ShowHidden = true; } catch { }

                // 记录已有的书签及其范围类型（是否包裹了外层的括号）
                Dictionary<string, bool> bookmarkCoverage = new Dictionary<string, bool>();
                foreach (Word.Bookmark bm in cellRightExisting.Range.Bookmarks)
                {
                    // 判断书签是否从单元格最左侧或者更左侧开始（包含了左括号）
                    bool coversParentheses = (bm.Range.Start <= rightRange.Start);
                    bookmarkCoverage[bm.Name] = coversParentheses;
                }

                try { doc.Bookmarks.ShowHidden = oldShowHidden; } catch { }

                // 删除旧的括号和编号
                rightRange.Text = "()";

                Word.Range rngNumExisting = existingTable.Cell(1, 3).Range;
                rngNumExisting.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                rngNumExisting.Move(Word.WdUnits.wdCharacter, 1);
                int bmStartPosExisting = rngNumExisting.Start;
                Word.Range bmRngExisting;

                if (numberStyle == 1)
                {
                    Word.Field field = doc.Fields.Add(rngNumExisting, Word.WdFieldType.wdFieldEmpty, @"SEQ 公式 \* ARABIC", false);
                    Word.Range lastRng = field.Result.Duplicate;
                    lastRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    lastRng.Move(Word.WdUnits.wdCharacter, 1);
                    bmRngExisting = doc.Range(bmStartPosExisting, lastRng.Start);
                }
                else
                {
                    Word.Range sepRng;
                    if (hasAutoNum)
                    {
                        Word.Field styleField = doc.Fields.Add(rngNumExisting, Word.WdFieldType.wdFieldEmpty, @"STYLEREF ""标题 1"" \s", false);
                        sepRng = styleField.Result.Duplicate;
                        sepRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        sepRng.Move(Word.WdUnits.wdCharacter, 1);
                    }
                    else
                    {
                        rngNumExisting.Text = chapterCount.ToString();
                        sepRng = rngNumExisting.Duplicate;
                        sepRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                    sepRng.Text = (numberStyle == 2) ? "." : "-";
                    Word.Range seqRng = sepRng.Duplicate;
                    seqRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Field seqField = doc.Fields.Add(seqRng, Word.WdFieldType.wdFieldEmpty, @"SEQ 公式 \* ARABIC \s 1", false);

                    Word.Range lastRng = seqField.Result.Duplicate;
                    lastRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    lastRng.Move(Word.WdUnits.wdCharacter, 1);
                    bmRngExisting = doc.Range(bmStartPosExisting, lastRng.Start);
                }

                if (bookmarkCoverage.Count > 0)
                {
                    Word.Range outerRng = existingTable.Cell(1, 3).Range;
                    outerRng.End -= 1; // 整个连带外层括号的范围

                    foreach (var kvp in bookmarkCoverage)
                    {
                        if (kvp.Value)
                        {
                            // 还原包裹了括号的整段书签（如自动交叉引用生成的 _Ref 书签）
                            doc.Bookmarks.Add(kvp.Key, outerRng);
                        }
                        else
                        {
                            // 还原仅包裹内部域代码的书签（如原生 OLE_LINK 书签）
                            doc.Bookmarks.Add(kvp.Key, bmRngExisting);
                        }
                    }
                }
                else
                {
                    string defaultBookmarkName = "OLE_LINK" + Guid.NewGuid().ToString("N").Substring(0, 8); 
                    doc.Bookmarks.Add(defaultBookmarkName, bmRngExisting);
                }

                cellRightExisting.Range.Font.Italic = 0; // 取消斜体
                return;
            }

            mathToWrap.Range.Select();
            app.Selection.Cut(); // 剪切现有的公式

            Word.PageSetup pageSetup = selection.PageSetup;
            // 右两侧专用于编号和占位的列宽
            float centerWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin - (sideWidth * 2);

            Word.Table table = doc.Tables.Add(selection.Range, 1, 3);
            table.Borders.Enable = 0;
            table.Rows[1].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.AllowAutoFit = false;
            table.Rows.LeftIndent = 0f;

            table.Cell(1, 1).Width = sideWidth;
            table.Cell(1, 2).Width = centerWidth;
            table.Cell(1, 3).Width = sideWidth;

            table.Cell(1, 1).LeftPadding = 0f;

            Word.Cell cellCenter = table.Cell(1, 2);
            cellCenter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            Word.Range rngCenter = cellCenter.Range;
            rngCenter.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            rngCenter.Select();
            app.Selection.Paste(); // 把刚刚剪切的公式粘贴回中间单元格

            Word.OMath rePastedMath = cellCenter.Range.OMaths[1];
            rePastedMath.Type = Word.WdOMathType.wdOMathDisplay;

            Word.Cell cellRight = table.Cell(1, 3);
            cellRight.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            Word.Range rngRight = cellRight.Range;
            rngRight.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            rngRight.Text = "()";

            Word.Range rngNum = table.Cell(1, 3).Range;
            rngNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            rngNum.Move(Word.WdUnits.wdCharacter, 1);

            int bmStartPos = rngNum.Start;
            Word.Range bmRng;

            if (numberStyle == 1)
            {
                Word.Field field = doc.Fields.Add(rngNum, Word.WdFieldType.wdFieldEmpty, @"SEQ 公式 \* ARABIC", false);
                Word.Range lastRng = field.Result.Duplicate;
                lastRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                lastRng.Move(Word.WdUnits.wdCharacter, 1);
                bmRng = doc.Range(bmStartPos, lastRng.Start);
            }
            else
            {
                Word.Range sepRng;
                if (hasAutoNum)
                {
                    Word.Field styleField = doc.Fields.Add(rngNum, Word.WdFieldType.wdFieldEmpty, @"STYLEREF ""标题 1"" \s", false);
                    sepRng = styleField.Result.Duplicate;
                    sepRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    sepRng.Move(Word.WdUnits.wdCharacter, 1);
                }
                else
                {
                    rngNum.Text = chapterCount.ToString();
                    sepRng = rngNum.Duplicate;
                    sepRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                sepRng.Text = (numberStyle == 2) ? "." : "-";
                Word.Range seqRng = sepRng.Duplicate;
                seqRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.Field seqField = doc.Fields.Add(seqRng, Word.WdFieldType.wdFieldEmpty, @"SEQ 公式 \* ARABIC \s 1", false);

                Word.Range lastRng = seqField.Result.Duplicate;
                lastRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                lastRng.Move(Word.WdUnits.wdCharacter, 1);
                bmRng = doc.Range(bmStartPos, lastRng.Start);
            }

            string bookmarkName = "OLE_LINK" + Guid.NewGuid().ToString("N").Substring(0, 8); 
            doc.Bookmarks.Add(bookmarkName, bmRng);

            cellRight.RightPadding = 0f;
            cellRight.Range.ParagraphFormat.RightIndent = 0f;
            cellRight.Range.ParagraphFormat.CharacterUnitRightIndent = 0f;
            cellRight.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            cellRight.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            cellRight.Range.Font.Italic = 0; // 取消斜体

            // 处理由于剪切并创建表格可能导致后续回车换行带有公式斜体格式的问题
            Word.Range afterTable = table.Range.Duplicate;
            afterTable.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterTable.Font.Italic = 0;

            Word.Range inputRng = rePastedMath.Range;
            inputRng.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            inputRng.Select();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
