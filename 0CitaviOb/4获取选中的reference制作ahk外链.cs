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

using System.IO;

/// <summary>
/// Citavi宏：获取选中的Reference文献信息并生成带跳转链接的引用
/// 
/// 功能概述：
/// 此宏用于快速获取当前在Citavi文献编辑器中选中的文献条目（Reference）信息，并将其格式化为包含引用文本和跳转链接的复合字符串，然后复制到剪贴板。
/// 
/// 生成的引用格式如下：
/// (第一作者 年份 – 分组名 标题前20个字符) [ahklink](ahk://citavi/goto?type=Ref&id=文献GUID&project=项目标识符&projectType=项目类型)
/// 
/// 工作流程：
/// 1.  获取当前文献：确保当前工作区是文献编辑器，并获取用户当前选中的文献条目。
/// 2.  提取并格式化信息：
///     - 作者：如果有多位作者，则自动格式化为 "第一作者 et al."。
///     - 年份：提取文献的发表年份。
///     - 分组：获取文献所属的第一个分组名称。
///     - 标题：取标题的前20个字符。
///     - 项目信息：获取当前Citavi项目的类型和标识符（路径或名称）。
/// 3.  组合字符串：
///     - 第一部分：生成学术引用格式的文本
///     - 第二部分：生成AutoHotkey跳转链接，始终包含 project 和 projectType 参数
/// 4.  复制到剪贴板：将生成的复合字符串复制到系统剪贴板。
/// 使用场景：
/// 当你在Obsidian或其他笔记软件中需要引用Citavi中的某篇文献，并希望后续能从笔记直接跳转回Citavi查看时，只需在Citavi中选中该文献，运行此宏，然后粘贴即可。
/// 它是构建"文献-笔记"双向链接体系中，从Citavi侧生成可跳转引用链接的便捷工具。
/// </summary>

public static class CitaviMacro
{
    public static void Main() 
    {
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

        mainForm.ActiveWorkspace = MainFormWorkspace.ReferenceEditor;
        Reference reference = mainForm.ActiveReference;
        string obsidianLink = FormatReferenceCitation(reference);
        Clipboard.SetText(obsidianLink);
        DebugMacro.WriteLine(obsidianLink);
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
    /// 根据文献信息生成包含引用文本和跳转链接的复合格式字符串。
    /// </summary>
    /// <param name="reference">要处理的文献条目对象。</param>
    /// <returns>格式化后的复合字符串。</returns>
    public static string FormatReferenceCitation(Reference reference)
    {
        if (reference == null)
        {
            return "Reference object is null.";
        }

        // --- 第一部分：生成引用文本 ---
        string titleToUse = reference.Title;
        if (titleToUse.Length > 20)
        {
            titleToUse = titleToUse.Substring(0, 20);
        }

        string authorsText = reference.Authors.ToStringSafe();
        if (reference.Authors != null && reference.Authors.Count > 1)
        {
            Person firstAuthor = reference.Authors.FirstOrDefault();
            if (firstAuthor != null && !string.IsNullOrEmpty(firstAuthor.LastName))
            {
                authorsText = firstAuthor.LastName + " et al.";
            }
            else
            {
                authorsText = reference.Authors.ToStringSafe();
            }
        }

        string yearText = reference.YearResolved.ToString();
		
        string referenceGroupname = "";
        if (reference.Groups.FirstOrDefault() != null)
        {
            referenceGroupname = reference.Groups.FirstOrDefault().FullName;
        }

        // 组合引用文本部分
        string citationText = string.Format("({0} {1} – {2} {3})",
            authorsText,
            yearText,
            referenceGroupname,
            titleToUse
        );

        // --- 第二部分：生成ahk跳转链接 ---
        string referenceId = reference.Id.ToStringSafe();
        
        // --- C# 4.6.1 兼容的写法 ---
        // 1. 先声明 out 变量
        string projectIdentifier;
        string projectType;

        // 2. 再调用方法并传递变量
        GetProjectInfo(Program.ActiveProjectShell.Project, out projectIdentifier, out projectType);

        // 构建URL，始终包含 project 和 projectType 参数
        string ahkUrl = string.Format("ahk://citavi/goto?type=Ref&id={0}&project={1}&projectType={2}", 
            referenceId, 
            projectIdentifier.Replace(" ","%20"), 
            projectType);
		
        // 将ahk链接包装在Obsidian的Markdown链接格式中，链接文本设为 "ahklink"
        string obsidianLink = string.Format("[ahklink]({0})", ahkUrl);

        // --- 第三部分：将两部分拼接在一起 ---
        string finalOutput = citationText + " " + obsidianLink;

        return finalOutput;
    }
}