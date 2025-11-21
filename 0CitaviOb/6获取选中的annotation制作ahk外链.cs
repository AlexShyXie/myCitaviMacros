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
 *                           Citavi PDF注释链接生成宏
 *
 * =================================================================================================
 *
 * 功能概述:
 * --------
 * 本宏用于获取用户在Citavi PDF预览界面（无论是右侧面板还是全屏模式）中选中的
 * PDF注释（Annotation），并生成一个包含跳转链接的复合字符串。该链接采用AutoHotkey
 * 协议，配合外部脚本可实现从笔记软件直接跳转回Citavi中对应的PDF注释位置。
 *
 * 工作流程:
 * --------
 * 1.  智能检测并获取用户当前选中的PDF注释对象，兼容右侧面板和全屏预览两种模式。
 * 2.  通过反射机制深入Citavi的UI组件，确保能稳定地捕获到选中的注释。
 * 3.  尝试从该注释对象反向查找其关联的知识条目（KnowledgeItem）。
 * 4.  提取关联知识条目的核心陈述（CoreStatement）作为链接的显示文本。
 *     - 如果找不到关联的知识条目，则使用默认文本"跳转到Annot"。
 * 5.  获取当前Citavi项目的文件路径，用于构建跳转链接。
 * 6.  生成AutoHotkey协议链接，格式为：
 *     ahk://citavi/goto?type=Annot&id=注释ID&project=项目路径
 * 7.  将显示文本和跳转链接组合成Obsidian Markdown格式的字符串。
 * 8.  将生成的字符串复制到系统剪贴板，并在宏控制台输出调试信息。

 * 使用场景:
 * --------
 * 当你在阅读PDF并做了高亮或批注后，希望在Obsidian或其他笔记软件中引用这个
 * 具体的标注位置时，只需在Citavi中选中该注释，运行此宏，然后粘贴即可。
 * 生成的链接配合相应的AutoHotkey脚本，可以让你在笔记中点击链接直接定位到
 * Citavi的PDF页面和对应的注释，实现精确定位和知识回溯。
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
		// 1. 先声明 out 变量
        string projectIdentifier;
        string projectType;

        // 2. 再调用方法并传递变量
        GetProjectInfo(Program.ActiveProjectShell.Project, out projectIdentifier, out projectType);
		
		string ahkUrl = string.Format("ahk://citavi/goto?type=Annot&id={0}&project={1}&projectType={2}", selectedAnnotation.Id.ToStringSafe(), projectIdentifier.Replace(" ", "%20"),projectType);
		// 将ahk链接包装在Obsidian的Markdown链接格式中，链接文本设为 "ahklink"
		string obsidianLink = string.Format("[ahklink]({0})", ahkUrl);
		string coreText = "";
		if (linkedKnowledge == null)
        {
			coreText = "跳转到Annot";
		}else
		{
			coreText = linkedKnowledge.CoreStatement;
		}
		string finalOutput = coreText + " " + obsidianLink;
		Clipboard.SetText(finalOutput);
		DebugMacro.WriteLine(finalOutput);

    }

	/// <summary>
    /// 获取项目的标识符（路径或名称）和类型。
    /// </summary>
    /// <param name="activeProject">当前活动的Citavi项目。</param>
    /// <param name="projectIdentifier">返回项目的标识符（本地项目为路径，服务器项目为名称）。</param>
    /// <param name="projectType">返回项目的类型（如 "DesktopSQLite" 或 "DesktopSqlServer"）。</param>
    public static void GetProjectInfo(Project activeProject, out string projectIdentifier, out string projectType)
    {
        projectIdentifier = String.Empty;
        projectType = activeProject.ProjectType.ToStringSafe();

        if (activeProject.DesktopProjectConfiguration.ProjectType == ProjectType.DesktopSQLite)
        {
            // 本地项目：标识符是文件路径
            projectIdentifier = activeProject.DesktopProjectConfiguration.SQLiteProjectInfo.FilePath;
        }
        else
        {
            // 服务器项目：标识符是项目名称
            projectIdentifier = activeProject.Name;
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