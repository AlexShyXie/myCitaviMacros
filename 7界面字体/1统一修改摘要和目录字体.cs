using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using SwissAcademic;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using SwissAcademic.Controls.WordProcessor;

// 修改第45行的字体大小，就可以改变Abstract和TableOfContent的预览大小了
// 在215行添加了设置加粗，这样整个改变是可以被保存的

//重要提示：此宏可能会导致 Citavi 在执行完成后崩溃
// (影响 Citavi 3.0.18 及更早版本，后续版本已修复)。
// 尽管如此，此宏所做的所有更改都会被保存。
// 如果 Citavi 假死，请使用任务管理器 (CTRL+ALT+DELETE) 结束挂起的 Citavi.exe 进程。
// 否则您将无法重新启动 Citavi。
// 宏编辑器的实现是初步且实验性的。
// Citavi 对象模型在未来的版本中可能会发生变化。

public static class CitaviMacro
{
    public static void Main()
    {
        if (Program.ProjectShells.Count == 0) return; //没有打开的项目
        if (IsBackupAvailable() == false) return; //用户想要先备份项目

        int counter = 0;

        try
        {
            CommonParagraphAttributes paragraphFormat = new CommonParagraphAttributes();

            //*******************************************************************************************************************//
            //注意：您可以根据需要调整以下内容：

            //------------- 范围 ----------------------------------------------------------------------------------------------//
            // 您是否想包含摘要、目录和评价字段？
            bool includeReferencePropertyFields = true; // 如果只想影响“真正的”知识条目，请设置为 false

            //------------------ 字体 -------------------------------------------------------------------------------------------//
            string fontName = "Microsoft YaHei UI"; //注意：字体名称
            int fontSize = 16; //注意：大小（点）

            //-------------------------------------------------------------------------------------------------------------------//

            //------------------ 段落对齐 ----------------------------------------------------------------------------//
            paragraphFormat.Alignment = Alignment.Left; //注意：左对齐
            //注意：或者：
            //paragraphFormat.Alignment = Alignment.Justify; //注意：两端对齐
            //注意：或者：
            //paragraphFormat.Alignment = Alignment.Center; //注意：居中
            //注意：或者：
            //paragraphFormat.Alignment = Alignment.Right; //注意：右对齐

            //-------------------------------------------------------------------------------------------------------------------//

            //------------------ 段落内行间距 ------------------------------------------------------------------//
            paragraphFormat.Spacing.LineSpacingType = LineSpacingType.One; //注意：单倍行距
            //注意：或者：
            //paragraphFormat.Spacing.LineSpacingType = LineSpacingType.OneAndHalf; //注意：1.5倍行距
            //注意：或者：
            //paragraphFormat.Spacing.LineSpacingType = LineSpacingType.Double; //注意：双倍行距
            //注意：或者：
            //paragraphFormat.Spacing.LineSpacingType = LineSpacingType.Multiple; //注意：多倍行距
            //float n = 3F; //注意：例如3倍行距，末尾总是带 F ！
            //paragraphFormat.Spacing.Line = Convert.ToInt32((n * 100F) - 100F);
            //注意：或者：
            //paragraphFormat.Spacing.LineSpacingType = LineSpacingType.Precise; //注意：固定值（点）
            //float pt = 12.5F; //注意：点数，末尾总是带 F，不要更改下一行
            //paragraphFormat.Spacing.Line = Convert.ToInt32(MeasurementUnit.Points.ConvertValue(pt, MeasurementUnitType.Twips));
            //注意：或者：
            //paragraphFormat.Spacing.LineSpacingType = LineSpacingType.Minimum; //注意：最小值（点）
            //float pt = 12.5F; //注意：点数，末尾总是带 F，不要更改下一行
            //paragraphFormat.Spacing.Line = Convert.ToInt32(MeasurementUnit.Points.ConvertValue(pt, MeasurementUnitType.Twips));

            //-------------------------------------------------------------------------------------------------------------------//

            //------------------ 段前/段后间距 -----------------------------------------------------------------//
            float ptAfter = 6F; //注意：点数，末尾总是带 F，不要更改下一行
            paragraphFormat.Spacing.After = Convert.ToInt32(MeasurementUnit.Points.ConvertValue(ptAfter, MeasurementUnitType.Twips));
            //注意：以及/或者：
            //float ptBefore = 12F; //注意：点数，末尾总是带 F，不要更改下一行
            //paragraphFormat.Spacing.Before = Convert.ToInt32(MeasurementUnit.Points.ConvertValue(ptBefore, MeasurementUnitType.Twips));

            //-------------------------------------------------------------------------------------------------------------------//

            //------------------ 段落缩进 -----------------------------------------------------------------------//
            //注意：重置 - 不要删除以下行 - 它们必须始终应用，否则会变得一团糟
            paragraphFormat.Indentation.Left = 0;
            paragraphFormat.Indentation.Right = 0;
            paragraphFormat.Indentation.FirstLine = 0;

            //注意：左缩进
            //float cmLeft = 0F; //注意：单位 cm，末尾总是带 F，不要编辑以下行，只需取消注释
            //paragraphFormat.Indentation.Left = Convert.ToInt32(MeasurementUnit.Centimeters.ConvertValue(cmLeft, MeasurementUnitType.Twips));

            //注意：右缩进
            //float cmRight = 0F; //注意：单位 cm，末尾总是带 F，不要编辑以下行，只需取消注释
            //paragraphFormat.Indentation.Right = Convert.ToInt32(MeasurementUnit.Centimeters.ConvertValue(cmRight, MeasurementUnitType.Twips));

            //注意：首行缩进
            //float cmFirstLine = 1F; //注意：单位 cm，末尾总是带 F，不要编辑以下行，只需取消注释
            //paragraphFormat.Indentation.IndentationType = IndentationType.FirstLine;
            //paragraphFormat.Indentation.FirstLine = Convert.ToInt32(MeasurementUnit.Centimeters.ConvertValue(cmFirstLine, MeasurementUnitType.Twips));

            //注意：或者：悬挂缩进
            //float cmHanging = 1F; //注意：单位 cm，末尾总是带 F，不要编辑以下行，只需取消注释
            //paragraphFormat.Indentation.IndentationType = IndentationType.Hanging;
            //paragraphFormat.Indentation.FirstLine = Convert.ToInt32(MeasurementUnit.Centimeters.ConvertValue(cmHanging, MeasurementUnitType.Twips));

            //-------------------------------------------------------------------------------------------------------------------//
            //*******************************************************************************************************************//

            //注意：请勿更改此行以下的任何内容

            //KnowledgeItem[] knowledgeItems = Program.ActiveProjectShell.Project.AllKnowledgeItems.ToArray();

            //引用活动项目 shell
            SwissAcademic.Citavi.Shell.ProjectShell activeShell = Program.ActiveProjectShell;
            //引用活动项目
            SwissAcademic.Citavi.Project activeProject = activeShell.Project;
            //引用主窗体
            SwissAcademic.Citavi.Shell.MainForm mainForm = activeShell.PrimaryMainForm;

            var references = mainForm.GetFilteredReferences();

            List<KnowledgeItem> knowledgeItemList = new List<KnowledgeItem>();
            foreach (Reference reference in references)
            {
                knowledgeItemList.AddRange(reference.Quotations);
                if (includeReferencePropertyFields)
                {
                    knowledgeItemList.Add(reference.Abstract);
                    knowledgeItemList.Add(reference.TableOfContents);
                    knowledgeItemList.Add(reference.Evaluation);
                }
            }

            foreach (KnowledgeItem thought in activeProject.Thoughts)
            {
                knowledgeItemList.Add(thought);
            }

            KnowledgeItem[] knowledgeItems = knowledgeItemList.ToArray();

            object tag = null;

            foreach (KnowledgeItem knowledgeItem in knowledgeItems)
            {
                if (string.IsNullOrEmpty(knowledgeItem.Text)) continue;
                if (knowledgeItem.KnowledgeItemType != KnowledgeItemType.Text && knowledgeItem.KnowledgeItemType != KnowledgeItemType.ReferenceProperty) continue;

                counter++;

                //首先我们实例化一个新的 RTF 窗体 ...
                //SwissAcademic.Citavi.Shell.RtfForm rtfForm = activeShell.ShowAbstractForm(reference);
                SwissAcademic.Citavi.Shell.ProjectShellForm rtfForm = null;
                switch (knowledgeItem.MirrorsReferencePropertyId)
                {
                    case ReferencePropertyId.Abstract:
                        rtfForm = activeShell.ShowAbstractForm(knowledgeItem.Reference);
                        break;
                    case ReferencePropertyId.TableOfContents:
                        rtfForm = activeShell.ShowTableOfContentsForm(knowledgeItem.Reference);
                        break;
                    case ReferencePropertyId.Evaluation:
                        rtfForm = activeShell.ShowEvaluationForm(knowledgeItem.Reference);
                        break;
                    default:
                        rtfForm = activeShell.ShowKnowledgeItemFormForExistingItem(mainForm, knowledgeItem);
                        break;
                }

                //rtfForm.PerformCommand("FormatRemoveReturnsAndTabs", tag);

                //... 然后通过反射获取其中的 WordProcessorControl 引用 ...
                Type t = rtfForm.GetType();
                FieldInfo fieldInfo = t.GetField("wordProcessor", BindingFlags.Instance | BindingFlags.NonPublic);
                if (fieldInfo == null) continue;

                WordProcessorControlEx wordProcessorControl = (WordProcessorControlEx)fieldInfo.GetValue(rtfForm);
                if (wordProcessorControl == null) continue;

                WordProcessorControl wordProcessor = wordProcessorControl.Editor;
                if (wordProcessor == null) continue;

                //... 并执行格式化：
				// 1. 插入文本
				wordProcessor.InsertTerText("摘要", true);

				// 2. 使用内置查找功能选中刚刚插入的文字
				// 解释：不用去计算光标左移了几格，直接告诉编辑器“找到‘摘要’这两个字并选中它”
				// flags 参数参考：2 通常代表不区分大小写等，具体数值需参考 Citavi/SwissAcademic 文档或常量定义
				int foundPos = wordProcessor.TerSearchReplace2("摘要", "", 2, 0, 50);
				// 如果找到了 (返回值 >= 0)
				if (foundPos >= 0)
				{
				    // 接下来直接设置加粗
				    wordProcessor.Select(foundPos, 2);
					wordProcessor.Selection.Bold = true;
					//rtfForm.PerformCommand("Save", tag);
				}

				wordProcessor.SelectAll();
                if (!string.IsNullOrEmpty(fontName))
                    wordProcessor.Selection.FontName = fontName;
                if (fontSize != 0)
                    wordProcessor.Selection.FontSize = Convert.ToInt32(SwissAcademic.MeasurementUnit.Points.ConvertValue(fontSize, SwissAcademic.MeasurementUnitType.Twips));

                // 添加这一行来设置加粗
                //wordProcessor.Selection.Bold = false;
                wordProcessor.Selection.SetCommonParagraphAttributes(paragraphFormat);
                rtfForm.PerformCommand("Save", tag);
                rtfForm.Close();
            }
        } //end try
        catch (Exception exception)
        {
            MessageBox.Show(exception.ToString());
        }
        finally
        {
            MessageBox.Show(string.Format("宏执行完毕。\r\n共进行了 {0} 处更改。", counter.ToString()), "Citavi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        } //end finally
    } //end main()

    private static bool IsBackupAvailable()
    {
        string warning = String.Concat(
            "重要提示：此宏将对您的项目进行不可逆的更改。",
            "\r\n\r\n",
            "在运行此宏之前，请确保您拥有当前项目的备份。",
            "\r\n",
            "如果您不确定，请单击“取消”，然后在 Citavi 主窗口的“文件”菜单上，单击“创建备份”。",
            "\r\n\r\n",
            "您想继续吗？"
            );

        return (MessageBox.Show(warning, "Citavi", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.OK);
    } //end IsBackupAvailable()
}