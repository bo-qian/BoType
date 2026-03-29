using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace BoType
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            int defaultStyle = Globals.ThisAddIn.LoadDefaultNumberStyle();
            this.dropDown1.SelectedItemIndex = (defaultStyle >= 0 && defaultStyle < 4) ? defaultStyle : 1;

            float defaultWidth = Globals.ThisAddIn.LoadDefaultSideWidth();
            this.editBoxWidth.Text = defaultWidth.ToString();
        }

        private float GetSideWidth()
        {
            if (float.TryParse(this.editBoxWidth.Text, out float width) && width > 0)
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
                System.Windows.Forms.MessageBox.Show("请选择一个带有编号的样式（非无）。", "BoType - 提示");
                return;
            }
            Globals.ThisAddIn.WrapSelectedEquation(styleIndex, GetSideWidth());
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            int styleIndex = this.dropDown1.SelectedItemIndex;
            if (styleIndex < 0) styleIndex = 1;
            Globals.ThisAddIn.SaveDefaultSettings(styleIndex, GetSideWidth());
            System.Windows.Forms.MessageBox.Show("已将所选编号样式和占位宽度设为开启软件时的默认样式。", "BoType - 提示");
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                // 打开 Word 内置的“交叉引用”对话框
                Word.Dialog dialog = app.Dialogs[Word.WdWordDialog.wdDialogInsertCrossReference];

                // 尝试将默认引用类型设置为 "公式"
                try
                {
                    dynamic dynDialog = dialog;
                    dynDialog.ReferenceType = "公式";
                    // 默认勾选“插入为超链接” (1代表勾选)
                    dynDialog.InsertAsHyperlink = 1;
                }
                catch
                {
                    // 忽略设置属性时的异常，以防有些语言版本不匹配
                }

                dialog.Show();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("无法打开交叉引用对话框: " + ex.Message, "BoType - 错误");
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

                    System.Windows.Forms.MessageBox.Show("引用更新完成！", "BoType - 提示");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("更新失败: " + ex.Message, "BoType - 错误");
            }
        }

        private void buttonGithub_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://github.com/bo-qian/BoType");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("无法打开链接: " + ex.Message, "BoType - 错误");
            }
        }
    }
}
