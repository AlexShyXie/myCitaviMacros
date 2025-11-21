// autoref "SwissAcademic.Pdf.dll"

using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System.Diagnostics;
using SwissAcademic.Pdf;
/*
 * =================================================================================================
 *
 *                               Citavi 知识条目链接生成宏
 *
 * =================================================================================================
 *
 * 功能概述:
 * --------
 * 本宏用于获取用户在 Citavi 中选中的知识条目（KnowledgeItem），并生成包含该条目
 * 核心陈述和跳转链接的复合字符串。生成的链接采用 AutoHotkey 协议，配合外部脚本
 * 可实现从笔记软件直接跳转回 Citavi 中对应的知识条目。
 *
 * 工作流程:
 * --------
 * 1.  检测当前活动工作区，支持从文献编辑器（ReferenceEditor）或知识组织器
 *     （KnowledgeOrganizer）中获取选中的知识条目。
 * 2.  验证用户是否精确选择了一个知识条目，确保操作的准确性。
 * 3.  提取知识条目的核心陈述（CoreStatement）作为引用文本。
 * 4.  获取当前 Citavi 项目的文件路径，用于构建跳转链接。
 * 5.  生成 AutoHotkey 协议链接，格式为：
 *     ahk://citavi/goto?type=Know&id=知识条目ID&project=项目路径
 * 6.  将核心陈述和跳转链接组合成 Obsidian Markdown 格式的字符串。
 * 7.  将生成的字符串复制到系统剪贴板，并在宏控制台输出调试信息。
 *
 * 关键技术点:
 * ----------
 * - 工作区适配：智能识别当前活动的工作区类型，支持多种获取选中项的方法。
 * - 项目路径处理：自动获取本地 SQLite 项目的完整路径，并处理路径格式转换。
 * - 链接构建：使用标准化的 AutoHotkey 协议格式，确保跳转功能的可靠性。
 * - 错误处理：包含选中项验证和项目类型检查，提供友好的错误提示。
 *
 * 使用场景:
 * --------
 * 当你在 Obsidian 或其他笔记软件中需要引用 Citavi 中的某个知识条目时，
 * 只需在 Citavi 中选中该条目，运行此宏，然后粘贴即可。生成的链接配合
 * 相应的 AutoHotkey 脚本，可以让你在笔记中点击链接直接跳转到 Citavi 中
 * 对应的知识条目，实现知识的快速定位和回溯。
 *
 * =================================================================================================
 */



public static class CitaviMacro
{
	public static void Main()
	{
		Project project = Program.ActiveProjectShell.Project;
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
	
		List<KnowledgeItem> selectedKnowledgeItems = null;

		// 根据当前活动的工作区类型选择合适的方法
		if (mainForm.ActiveWorkspace.ToString() == "ReferenceEditor")
		{
			selectedKnowledgeItems = mainForm.GetSelectedQuotations();
		}
		else if (mainForm.ActiveWorkspace.ToString() == "KnowledgeOrganizer")
		{
			selectedKnowledgeItems = mainForm.GetSelectedKnowledgeItems();
		}
		List<Location> selectedLocations = mainForm.GetSelectedElectronicLocations();
		
		// 1. 获取用户选中的 KnowledgeItem 和 Location,检查是否只选中了一个引文
		if (selectedKnowledgeItems == null || selectedKnowledgeItems.Count != 1)
		{
			MessageBox.Show("请先在引文列表中精确地选择 **一个** 引文。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
			return; // 停止执行宏
		}
		
		//Location location = selectedLocations[0];
		KnowledgeItem knowledgeItem = selectedKnowledgeItems[0];

		
		// 1. 先声明 out 变量
        string projectIdentifier;
        string projectType;

        // 2. 再调用方法并传递变量
        GetProjectInfo(Program.ActiveProjectShell.Project, out projectIdentifier, out projectType);

		string ahkUrl = string.Format("ahk://citavi/goto?type=Know&id={0}&project={1}&projectType={2}", knowledgeItem.Id.ToStringSafe(), projectIdentifier.Replace(" ", "%20"),projectType);
		
		// 将ahk链接包装在Obsidian的Markdown链接格式中，链接文本设为 "ahklink"
		string obsidianLink = string.Format("[ahklink]({0})", ahkUrl);
		string finalOutput = knowledgeItem.CoreStatement + " " + obsidianLink;
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

}