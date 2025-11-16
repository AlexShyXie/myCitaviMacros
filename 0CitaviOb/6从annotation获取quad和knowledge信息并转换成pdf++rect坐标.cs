// autoref "SwissAcademic.Pdf.dll"

using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Collections;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System.Diagnostics;
using SwissAcademic.Pdf;
using SwissAcademic.Citavi.Shell.Controls.Preview; // 显式引用

/*
 * =================================================================================================
 *
 *                     Citavi to Obsidian 精确链接生成宏 (兼容C# 4.6.1版)
 *
 * =================================================================================================
 *
 * 功能概述:
 * --------
 * 本宏是之前两个版本的结合与优化。它直接从用户在Citavi PDF预览中选中的注释
 * 出发，提取其精确的几何坐标，并反向查找关联的知识条目以获取文本内容。
 * 最终生成一个包含精确坐标、页码、引文文本和知识ID的Obsidian `rect` 链接。
 *
 * 工作流程:
 * --------
 * 1.  直接获取用户在PDF预览中选中的 `Annotation` 对象。
 * 2.  从该 `Annotation` 对象中提取 `Quads` 集合，计算总包围盒。
 * 3.  从 `Annotation` 反向查找其关联的 `KnowledgeItem`。
 * 4.  获取PDF文件的本地路径，并构建Obsidian库内的相对路径。
 * 5.  将所有信息（坐标、页码、文件名、引文文本、知识ID）组合成最终的Markdown链接。
 * 6.  将链接复制到剪贴板。
 *
 * =================================================================================================
 */

public static class CitaviMacro
{
    public static void Main()
    {
        // 1. 直接获取选中的 Annotation
        Annotation selectedAnnotation = TryGetSelectedAnnotation();

        if (selectedAnnotation == null)
        {
            MessageBox.Show("未能获取到选中的注释。\n\n请确保你在PDF预览中选中了一个高亮或注释。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 2. 从 Annotation 反向查找 KnowledgeItem
        KnowledgeItem linkedKnowledge = GetKnowledgeFromAnnotation(selectedAnnotation);
		string obsidianLink = "";
        if (linkedKnowledge == null)
        {
			//DebugMacro.WriteLine("选中的注释未关联到任何知识条目，无法生成包含引文内容的链接。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            // 即使没有关联知识，我们也可以生成一个只有坐标的链接
			obsidianLink = ConvertAnnotationToObsidianLink(selectedAnnotation, linkedKnowledge);
			DebugMacro.WriteLine("--- 没有关联知识，生成只有坐标的链接，生成的 Obsidian 链接 ---");
			DebugMacro.WriteLine(obsidianLink);
			DebugMacro.WriteLine("---------------------------");

			Clipboard.SetText(obsidianLink);
			MessageBox.Show(string.Format("没有关联知识，生成只有坐标的链接，链接已复制到剪贴板！\n\n{0}", obsidianLink), "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
		else
		{
			// 3. 调用核心转换函数，生成链接
			obsidianLink = ConvertAnnotationToObsidianLink(selectedAnnotation, linkedKnowledge);
			obsidianLink = string.Format("{0} {1}",linkedKnowledge.CoreStatement,obsidianLink);
			if (!string.IsNullOrEmpty(obsidianLink))
			{
				DebugMacro.WriteLine("--- 生成的 Obsidian 链接 ---");
				DebugMacro.WriteLine(obsidianLink);
				DebugMacro.WriteLine("---------------------------");

				Clipboard.SetText(obsidianLink);
				//MessageBox.Show(string.Format("链接已复制到剪贴板！\n\n{0}", obsidianLink), "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			else
			{
				MessageBox.Show("生成链接失败。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			
		}

    }

    /// <summary>
    /// 尝试从右侧面板或全屏预览中获取用户选中的第一个 Annotation 对象。
    /// </summary>
    /// <returns>返回选中的 Annotation 对象，如果未找到则返回 null。</returns>
    public static Annotation TryGetSelectedAnnotation()
    {
        DebugMacro.WriteLine("尝试从右侧面板获取选中的注释...");
        PreviewControl mainPreviewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
        Annotation annotationFromMainPanel = GetAnnotationFromPreviewControl(mainPreviewControl);
        if (annotationFromMainPanel != null)
        {
            DebugMacro.WriteLine("成功从右侧面板获取到注释。");
            return annotationFromMainPanel;
        }
        DebugMacro.WriteLine("右侧面板未找到选中的注释。");

        DebugMacro.WriteLine("尝试从全屏预览窗口获取选中的注释...");
        Annotation annotationFromFullScreen = GetAnnotationFromFullScreenPreview();
        if (annotationFromFullScreen != null)
        {
            DebugMacro.WriteLine("成功从全屏预览窗口获取到注释。");
            return annotationFromFullScreen;
        }
        DebugMacro.WriteLine("全屏预览窗口也未找到选中的注释。");

        return null;
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

    /// <summary>
    /// 将一个 Citavi Annotation 对象（及其关联的 KnowledgeItem）转换为 Obsidian 的 rect 链接格式。
    /// </summary>
    /// <param name="annotation">要转换的 Citavi Annotation 对象。</param>
    /// <param name="knowledgeItem">关联的 KnowledgeItem 对象，可以为 null。</param>
    /// <returns>返回 Obsidian 链接字符串，如果转换失败则返回 null。</returns>
    public static string ConvertAnnotationToObsidianLink(Annotation annotation, KnowledgeItem knowledgeItem)
    {
        string ObsidianPdfBasePath = "11_影像学习/9_书籍PDF"; //这里设置你的Obsidian书库路径

        // 1. 基本校验
        if (annotation == null) return null;

        // 2. 获取PDF文件信息
        Location location = annotation.Location;
        if (location == null) return null;
        if (location.Address == null) return null;

		string locationGroupname ="";
		if (location.Reference.Groups.FirstOrDefault() != null)
		{
			locationGroupname = location.Reference.Groups.FirstOrDefault().FullName;
		}
		
        string fullFilePath = location.Address.Resolve().LocalPath;
        if (string.IsNullOrEmpty(fullFilePath) || !File.Exists(fullFilePath))
        {
            DebugMacro.WriteLine("无法获取有效的本地文件路径或文件不存在。解析后的路径: " + fullFilePath);
            return null;
        }
        string pdfFileName = Path.GetFileName(fullFilePath);
        string parentDirectoryName = Path.GetFileName(Path.GetDirectoryName(fullFilePath));

        // 3. 构建Obsidian库内路径
        string encodedFileName = pdfFileName.Replace(" ", "%20");
        string encodedParentDirectoryName = parentDirectoryName.Replace(" ", "%20");
        string obsidianRelativePath = string.Format("{0}/{1}/{2}", ObsidianPdfBasePath, encodedParentDirectoryName, encodedFileName);

        // 4. 计算总包围盒
        double minX = double.MaxValue;
        double minY = double.MaxValue;
        double maxX = double.MinValue;
        double maxY = double.MinValue;
        bool hasValidQuads = false;

        if (annotation.Quads != null)
        {
            foreach (var quad in annotation.Quads)
            {
                if (quad.IsContainer) continue;
                hasValidQuads = true;
                if (quad.MinX < minX) minX = quad.MinX;
                if (quad.MinY < minY) minY = quad.MinY;
                if (quad.MaxX > maxX) maxX = quad.MaxX;
                if (quad.MaxY > maxY) maxY = quad.MaxY;
            }
        }

        if (!hasValidQuads) return null;

        // 5. 应用像素填充
        //int horizontalPadding = 2;
        //int verticalPadding = 2;
        minX -= 2;
        maxX += 4;
        minY -= 2;
        maxY += 2;

        // 6. 转换坐标和页面索引
        int obsX1 = (int)Math.Round(minX);
        int obsY1 = (int)Math.Round(minY);
        int obsX2 = (int)Math.Round(maxX);
        int obsY2 = (int)Math.Round(maxY);
        int pageIndex = annotation.Quads.First().PageIndex;

        // 7. 生成链接
        string linkText;
        string finalLink;

        if (knowledgeItem != null)
        {
            // 如果关联了知识条目，使用知识条目的核心文本作为链接文本
            linkText = string.Format("{0}, p.{1}", parentDirectoryName+" "+ pdfFileName, pageIndex);
            finalLink = string.Format("[{0}]({1}#page={2}&rect={3},{4},{5},{6}&color=yellow) KnowID：{7} AnnotID：{8}",
                linkText, obsidianRelativePath, pageIndex, obsX1, obsY1, obsX2, obsY2, knowledgeItem.Id.ToStringSafe(), annotation.Id.ToStringSafe());
        }
        else
        {
            // 如果没有关联知识条目，使用文件名作为链接文本
            linkText = string.Format("{0}, p.{1}", parentDirectoryName+" "+ pdfFileName, pageIndex);
            finalLink = string.Format("[{0}{1}]({2}#page={3}&rect={4},{5},{6},{7}&color=yellow) AnnotID：{8}",
               locationGroupname, linkText, obsidianRelativePath, pageIndex, obsX1, obsY1, obsX2, obsY2, annotation.Id.ToStringSafe());
        }

        return finalLink;
    }

    // --- 以下是辅助方法，直接从你之前的宏中复制而来，无需改动 ---

    public static Annotation GetAnnotationFromPreviewControl(PreviewControl previewControl)
    {
        if (previewControl == null) return null;

        PropertyInfo pdfViewControlProperty = previewControl.GetType().GetProperty("PdfViewControl", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
        if (pdfViewControlProperty == null) return null;
        object pdfViewControlObject = pdfViewControlProperty.GetValue(previewControl);
        if (pdfViewControlObject == null) return null;

        PropertyInfo toolProperty = pdfViewControlObject.GetType().GetProperty("Tool");
        if (toolProperty == null) return null;
        object toolObject = toolProperty.GetValue(pdfViewControlObject);
        if (toolObject == null) return null;

        FieldInfo selectedContainersField = toolObject.GetType().GetField("SelectedAdornmentContainers", BindingFlags.Instance | BindingFlags.NonPublic);
        if (selectedContainersField == null) return null;
        object selectedContainersObject = selectedContainersField.GetValue(toolObject);
        if (selectedContainersObject == null) return null;
        
        IEnumerator enumerator = (selectedContainersObject as IEnumerable).GetEnumerator();
        
        var annotations = new List<object>();
        while (enumerator.MoveNext())
        {
            object container = enumerator.Current;
            if (container == null) continue;
            Type containerType = container.GetType();
            PropertyInfo annotationProperty = containerType.GetProperty("Annotation");
            if (annotationProperty == null) continue;
            object annotation = annotationProperty.GetValue(container);
            if (annotation != null)
            {
                annotations.Add(annotation);
            }
        }

        // 直接返回第一个找到的 Annotation
        return annotations.OfType<SwissAcademic.Citavi.Annotation>().FirstOrDefault();
    }

    public static Annotation GetAnnotationFromFullScreenPreview()
    {
        var projectShell = Program.ActiveProjectShell;
        if (projectShell == null) return null;

        var field = projectShell.GetType().GetField("_previewFullScreenForms", BindingFlags.Instance | BindingFlags.NonPublic);
        if (field == null) return null;
        object fullScreenFormsObject = field.GetValue(projectShell);
        if (fullScreenFormsObject == null) return null;

        PropertyInfo countProperty = fullScreenFormsObject.GetType().GetProperty("Count");
        if (countProperty == null) return null;
        int count = (int)countProperty.GetValue(fullScreenFormsObject);
        if (count == 0) return null;

        PropertyInfo indexerProperty = fullScreenFormsObject.GetType().GetProperty("Item");
        if (indexerProperty == null) return null;

        MainForm activeFullScreenForm = null;
        for (int i = 0; i < count; i++)
        {
            object formObject = indexerProperty.GetValue(fullScreenFormsObject, new object[] { i });
            MainForm form = formObject as MainForm;
            if (form != null && form.Visible)
            {
                activeFullScreenForm = form;
            }
        }
        
        if (activeFullScreenForm != null && activeFullScreenForm.PreviewControl != null)
        {
            return GetAnnotationFromPreviewControl(activeFullScreenForm.PreviewControl);
        }

        return null;
    }
}