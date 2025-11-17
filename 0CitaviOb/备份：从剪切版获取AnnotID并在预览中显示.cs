// autoref "SwissAcademic.Pdf.dll"

// autoref "SwissAcademic.Pdf.dll"

using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System.IO;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using System.Diagnostics;
using SwissAcademic.Citavi.Shell.Controls.SmartRepeaters;

// =================================================================================================
// Citavi 宏：通过剪贴板中的ID定位注释并根据关联知识项类型进行跳转 (重构版)
// =================================================================================================

public static class CitaviMacro
{
    /// <summary>
    /// 宏的主入口点。
    /// </summary>
    public static async Task Main() 
    {
        // --- 1. 初始化项目与主窗口引用 ---
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
        // --- 2. 从剪贴板读取ID并查找注释 ---
        List<Annotation> allAnnotations = project.AllAnnotations.ToList();
        List<Annotation> foundAnnotations = new List<Annotation>();

        using (var reader = new StringReader(Clipboard.GetText()))
        {
            for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
            {
                Annotation thisAnnotation = allAnnotations.Where(k => k.Id.ToString() == line.Trim()).FirstOrDefault();
                if (thisAnnotation == null) continue;
                foundAnnotations.Add(thisAnnotation);
            }
        }
		
        if(foundAnnotations.Count == 0)
        {
            MessageBox.Show("剪贴板中未找到有效的注释ID。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // --- 3. 获取第一个找到的注释并准备跳转 ---
        Annotation targetAnnotation = foundAnnotations[0];
        Location targetLocation = targetAnnotation.Location;

        // 4. 设置活动参考文献
        mainForm.ActiveWorkspace = MainFormWorkspace.ReferenceEditor;
        mainForm.ActiveReference = targetLocation.Reference;
		// 4.1. 确保中心面板可见
        if (mainForm.IsCenterPaneCollapsed)
        {
            mainForm.IsCenterPaneCollapsed = false;
        }
        mainForm.ActiveReferenceEditorTabPage = MainFormReferencesTabPage.Quotations;


		
        // --- 4.2. 判断注释是否关联了知识项 ---
        KnowledgeItem targetKnowledgeItem = GetKnowledgeFromAnnotation(targetAnnotation);

        // 情况一：注释没有关联知识项
        if (targetKnowledgeItem == null)
        {
            string warningMessage = "该Annot没有Knowledge，当前高亮的Knowledge不与Annot对应\n     Annot已在PDF中高亮，其他无问题";
            JumpToPdfAndHighlightAnnotation(mainForm, targetAnnotation, targetLocation, warningMessage);
        }
        // 情况二：注释关联了知识项
        else if(targetKnowledgeItem != null)
        {
            // 进一步判断知识项的类型
            if (targetKnowledgeItem.QuotationType != QuotationType.Highlight)
            {
                // 2b. 如果是其他类型（如摘要、引文、评论），高亮知识项并跳转到PDF
                await HighlightKnowledgeItemInRepeater(mainForm, targetKnowledgeItem);
            }
            else 
            {
                // 2a. 如果是高亮，跳转到PDF位置
                string warningMessage = "Annot是黄色高亮，对应的Knowledge是隐藏数据，无法高亮Knowledge\n      Annot已在PDF中高亮，其他无问题";
                JumpToPdfAndHighlightAnnotation(mainForm, targetAnnotation, targetLocation, warningMessage);
            }
        }
    }

    /// <summary>
    /// 强制刷新预览区，显示PDF内容，并精确跳转到指定注释。
    /// </summary>
    /// <param name="mainForm">Citavi主窗体</param>
    /// <param name="targetAnnotation">要跳转的目标注释</param>
    /// <param name="targetLocation">注释关联的位置</param>
    /// <param name="successWarningMessage">成功跳转后要显示的警告信息</param>
    private static void JumpToPdfAndHighlightAnnotation(MainForm mainForm, Annotation targetAnnotation, Location targetLocation, string successWarningMessage)
    {
        // --- 5. 强制刷新预览区，显示PDF内容 ---
        var previewControl = mainForm.PreviewControl;
        try
        {
            Type[] parameterTypes = new Type[] { typeof(Location), typeof(Reference), typeof(SwissAcademic.Citavi.PreviewBehaviour), typeof(bool) };
            var showLocationPreviewMethod = typeof(PreviewControl).GetMethod("ShowLocationPreview", BindingFlags.Public | BindingFlags.Instance, null, parameterTypes, null);
            
            if (showLocationPreviewMethod != null)
            {
                var previewBehaviourEnum = typeof(SwissAcademic.Citavi.PreviewBehaviour);
                var skipEntryPageValue = Enum.Parse(previewBehaviourEnum, "SkipEntryPage");
                object[] parameters = new object[] { targetLocation, targetLocation.Reference, skipEntryPageValue, true };
                
                showLocationPreviewMethod.Invoke(previewControl, parameters);
                Program.ActiveProjectShell.ShowMainForm();
            }
        }
        catch (Exception ex)
        {
            DebugMacro.WriteLine("调用预览时发生错误: " + ex.Message);
            MessageBox.Show("预览PDF时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- 6. 执行精确跳转 ---
        System.Threading.Thread.Sleep(1500); 

        PropertyInfo pdfViewControlProperty = previewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Instance);
        object pdfViewControlObject = null;
        if (pdfViewControlProperty != null)
        {
            pdfViewControlObject = pdfViewControlProperty.GetValue(previewControl);
        }

        if (pdfViewControlObject != null)
        {
            Type[] goToAnnotationParamTypes = new Type[] { typeof(Annotation), typeof(EntityLink) };
            MethodInfo goToAnnotationMethod = pdfViewControlObject.GetType().GetMethod("GoToAnnotation", goToAnnotationParamTypes);

            if (goToAnnotationMethod != null)
            {
                try
                {
                    object result = goToAnnotationMethod.Invoke(pdfViewControlObject, new object[] { targetAnnotation, null });
                    
                    if (result is bool && !(bool)result)
                    {
                        MessageBox.Show("PDF已加载，但未能跳转到指定注释。可能该注释在当前PDF中不可见。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        // 使用传入的自定义警告信息
                        MessageBox.Show(successWarningMessage, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("调用 GoToAnnotation 时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("在PdfViewControl中未找到 GoToAnnotation 方法。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        else
        {
            MessageBox.Show("无法获取PdfViewControl对象。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// 在Citavi界面中高亮指定的知识项，并跳转到其关联的PDF位置。
    /// </summary>
    /// <param name="mainForm">Citavi主窗体</param>
    /// <param name="targetKnowledgeItem">要高亮和跳转的知识项</param>
    private static async Task HighlightKnowledgeItemInRepeater(MainForm mainForm, KnowledgeItem targetKnowledgeItem)
    {
        QuotationSmartRepeater quotationSmartRepeater = null;
        
		if (Program.ActiveProjectShell.PrimaryMainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
        {
            quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;
        }
        else if (Program.ActiveProjectShell.PrimaryMainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
        {
            SmartRepeater<KnowledgeItem> KnowledgeItemSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("SmartRepeater", true).FirstOrDefault() as SmartRepeater<KnowledgeItem>;
            quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("knowledgeItemPreviewSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;

        }

		
        quotationSmartRepeater.SelectAndActivate(targetKnowledgeItem, true);

		// --- 5. 执行PDF跳转 ---
        // 检查目标知识条目是否确实关联了一个有效的地址（通常是PDF文件）。
        if (targetKnowledgeItem.Address != null)
        {
            // 调用核心方法，在PDF预览中异步跳转到知识条目对应的位置。
            await Program.ActiveProjectShell.PrimaryMainForm.PreviewControl.ShowPdfLinkAsync(mainForm, targetKnowledgeItem);
			Program.ActiveProjectShell.ShowMainForm();
        }
        else
        {
            // 如果知识条目没有关联地址（例如，它是一个纯文本的摘要），则提示用户。
            MessageBox.Show("知识条目 '{targetKnowledgeItem.CoreStatement}' 没有关联的PDF文件。", "无法跳转", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
		
		
    }
	
    /// <summary>
    /// 从一个 Annotation 对象反向查找其关联的 KnowledgeItem。
    /// </summary>
    /// <param name="annotation">要查询的 Annotation 对象。</param>
    /// <returns>返回关联的 KnowledgeItem，如果未找到则返回 null。</returns>
    public static KnowledgeItem GetKnowledgeFromAnnotation(Annotation annotation)
    {
        if (annotation == null || annotation.EntityLinks == null) return null;

        var targetLink = annotation.EntityLinks
            .Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication)
            .FirstOrDefault();

        if (targetLink != null && targetLink.Source is KnowledgeItem)
        {
            return (KnowledgeItem)targetLink.Source;
        }

        return null;
    }
}