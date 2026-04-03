using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace BoType
{
    public partial class BoTypeRibbon
    {
        private void BoTypeRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            int defaultStyle = Globals.ThisAddIn.LoadDefaultNumberStyle();
            this.dropDown1.SelectedItemIndex = (defaultStyle >= 0 && defaultStyle < 4) ? defaultStyle : 1;

            float defaultWidth = Globals.ThisAddIn.LoadDefaultSideWidth();
            this.comboBoxWidth.Text = defaultWidth.ToString() + " 磅";

            this.dropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            UpdateRefStyleOptions();

            try
            {
                Globals.ThisAddIn.Application.DocumentChange += Application_DocumentChange;
            }
            catch { }

            RefreshEquationList();
        }

        private void Application_DocumentChange()
        {
            RefreshEquationList();
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            UpdateRefStyleOptions();
        }

        private void UpdateRefStyleOptions()
        {
            int styleIndex = this.dropDown1.SelectedItemIndex;
            string ph = "1";
            if (styleIndex == 2) ph = "1.1";
            else if (styleIndex == 3) ph = "1-1";

            this.dropDownRefStyle.Items.Clear();
            var factory = Globals.Factory.GetRibbonFactory();

            var item0 = factory.CreateRibbonDropDownItem();
            item0.Label = $"({ph})";
            this.dropDownRefStyle.Items.Add(item0);

            var item1 = factory.CreateRibbonDropDownItem();
            item1.Label = $"{ph}";
            this.dropDownRefStyle.Items.Add(item1);

            var item2 = factory.CreateRibbonDropDownItem();
            item2.Label = $"Eq. {ph}";
            this.dropDownRefStyle.Items.Add(item2);

            var item3 = factory.CreateRibbonDropDownItem();
            item3.Label = $"Equation {ph}";
            this.dropDownRefStyle.Items.Add(item3);

            var item4 = factory.CreateRibbonDropDownItem();
            item4.Label = $"公式 {ph}";
            this.dropDownRefStyle.Items.Add(item4);

            int savedDefault = Globals.ThisAddIn.LoadDefaultRefStyle();
            if (savedDefault >= 0 && savedDefault < this.dropDownRefStyle.Items.Count)
            {
                this.dropDownRefStyle.SelectedItemIndex = savedDefault;
            }
            else
            {
                this.dropDownRefStyle.SelectedItemIndex = 0;
            }
        }

        private void buttonSetDefaultRefStyle_Click(object sender, RibbonControlEventArgs e)
        {
            int styleIndex = this.dropDownRefStyle.SelectedItemIndex;
            if (styleIndex < 0) styleIndex = 0;
            Globals.ThisAddIn.SaveDefaultRefStyle(styleIndex);
        }

        private float GetSideWidth()
        {
            string txt = this.comboBoxWidth.Text.Replace("磅", "").Replace("pt", "").Trim();
            if (float.TryParse(txt, out float width) && width > 0)
            {
                return width;
            }
            return 38.0f; // 默认值
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            int styleIndex = this.dropDown1.SelectedItemIndex;
            if (styleIndex < 0) styleIndex = 1; // 默认选择纯数字(1)
            Globals.ThisAddIn.InsertNumberedEquation(styleIndex, GetSideWidth());
            RefreshEquationList();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.InsertInlineEquation();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            int styleIndex = this.dropDown1.SelectedItemIndex;
            if (styleIndex <= 0) 
            {
                return;
            }
            Globals.ThisAddIn.WrapSelectedEquation(styleIndex, GetSideWidth());
            RefreshEquationList();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            int styleIndex = this.dropDown1.SelectedItemIndex;
            if (styleIndex < 0) styleIndex = 1;
            Globals.ThisAddIn.SaveDefaultSettings(styleIndex, GetSideWidth());
        }

        internal class EqInfo
        {
            public string BookmarkName { get; set; }
            public string DisplayText { get; set; }
            public int StartPos { get; set; }
        }

        private void RefreshEquationList()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            this.dropDownRefEq.Items.Clear();
            var factory = Globals.Factory.GetRibbonFactory();

            if (app.Documents.Count == 0) 
            {
                var item = factory.CreateRibbonDropDownItem();
                item.Label = "无";
                this.dropDownRefEq.Items.Add(item);
                return;
            }
            Word.Document doc;
            try { doc = app.ActiveDocument; } 
            catch
            {
                var item = factory.CreateRibbonDropDownItem();
                item.Label = "无";
                this.dropDownRefEq.Items.Add(item);
                return;
            }

            List<EqInfo> eqs = new List<EqInfo>();

            bool oldShowHidden = false;
            try { oldShowHidden = doc.Bookmarks.ShowHidden; doc.Bookmarks.ShowHidden = true; } catch { }

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table.Rows.Count != 1 || table.Columns.Count != 3) continue;
                    Word.Cell centerCell;
                    Word.Cell rightCell;
                    try
                    {
                        centerCell = table.Cell(1, 2);
                        rightCell = table.Cell(1, 3);
                    }
                    catch { continue; }

                    if (centerCell.Range.OMaths.Count == 0) continue;

                    string bmName = null;
                    string dispText = null;
                    int startPos = rightCell.Range.Start;

                    foreach (Word.Bookmark bm in rightCell.Range.Bookmarks)
                    {
                        if (bm.Name.StartsWith("OLE_LINK") || bm.Name.StartsWith("公式"))
                        {
                            bmName = bm.Name;
                            dispText = bm.Range.Text;
                            startPos = bm.Range.Start;
                            break;
                        }
                    }

                    if (!string.IsNullOrEmpty(bmName))
                    {
                        if (dispText != null) {
                            dispText = dispText.Replace("\r", "").Replace("\a", "").Trim();
                        }

                        eqs.Add(new EqInfo() { 
                            BookmarkName = bmName, 
                            DisplayText = $"({dispText})", 
                            StartPos = startPos 
                        });
                    }
                }

                eqs.Sort((a,b) => a.StartPos.CompareTo(b.StartPos));

                if (eqs.Count == 0)
                {
                    var item = factory.CreateRibbonDropDownItem();
                    item.Label = "无";
                    this.dropDownRefEq.Items.Add(item);
                    return;
                }

                foreach (var eq in eqs)
                {
                    Microsoft.Office.Tools.Ribbon.RibbonDropDownItem item = factory.CreateRibbonDropDownItem();
                    item.Label = eq.DisplayText;
                    item.Tag = eq.BookmarkName;
                    this.dropDownRefEq.Items.Add(item);
                }
            }
            finally
            {
                try { doc.Bookmarks.ShowHidden = oldShowHidden; } catch { }
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.dropDownRefEq.SelectedItem == null || this.dropDownRefEq.SelectedItem.Tag == null)
            {
                return;
            }

            string bookmarkName = this.dropDownRefEq.SelectedItem.Tag.ToString();
            int refStyleIndex = this.dropDownRefStyle.SelectedItemIndex;
            string prefix = "";
            bool withParens = false;

            if (refStyleIndex == 0) { prefix = ""; withParens = true; }
            else if (refStyleIndex == 1) { prefix = ""; withParens = false; }
            else if (refStyleIndex == 2) { prefix = "Eq. "; withParens = false; }
            else if (refStyleIndex == 3) { prefix = "Equation "; withParens = false; }
            else if (refStyleIndex == 4) { prefix = "公式 "; withParens = false; }

            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                app.ScreenUpdating = false;
                Word.Selection sel = app.Selection;
                sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                if (withParens) sel.TypeText("(");
                if (!string.IsNullOrEmpty(prefix)) sel.TypeText(prefix);

                // 插入引用域，如 REF OLE_LINK... \h
                Word.Field refField = app.ActiveDocument.Fields.Add(sel.Range, Word.WdFieldType.wdFieldEmpty, $@"REF {bookmarkName} \h", false);

                Word.Range afterField = refField.Result.Duplicate;
                afterField.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                afterField.Select();

                if (withParens) sel.TypeText(")");
            }
            catch (Exception)
            {
            }
            finally
            {
                try { Globals.ThisAddIn.Application.ScreenUpdating = true; } catch { }
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                if (app.Documents.Count > 0)
                {
                    Word.Document doc = app.ActiveDocument;

                    // 更新带有手动章节编号的公式
                    Globals.ThisAddIn.UpdateEquationChapterNumbers();

                    // 更新主文档中的所有域（包含 STYLEREF、SEQ 和 交叉引用）
                    doc.Fields.Update();

                    // 遍历所有的文本故事（包括页眉、页脚、文本框等）以确保彻底更新
                    foreach (Word.Range storyRange in doc.StoryRanges)
                    {
                        Word.Range currentRange = storyRange;
                        while (currentRange != null)
                        {
                            currentRange.Fields.Update();
                            currentRange = currentRange.NextStoryRange;
                        }
                    }

                    RefreshEquationList();
                }
            }
            catch (Exception)
            {
            }
        }

        private void buttonGithub_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://github.com/bo-qian/BoType");
            }
            catch (Exception)
            {
            }
        }
    }
}
