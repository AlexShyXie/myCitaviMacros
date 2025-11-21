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
/// Citavi宏：获取选中的Reference文献信息并生成格式化引用
/// 
/// 功能概述：
/// 此宏用于快速获取当前在Citavi文献编辑器中选中的文献条目（Reference）信息，并将其格式化为一个简洁、标准化的引用字符串，然后复制到剪贴板。
/// 
/// 生成的引用格式如下：
/// (第一作者 et al. 年份 – 标题前10个字符) RefID：文献GUID
/// 
/// 工作流程：
/// 1.  获取当前文献：确保当前工作区是文献编辑器，并获取用户当前选中的文献条目。
/// 2.  提取并格式化信息：
///     - 作者：如果有多位作者，则自动格式化为 "第一作者 et al."。
///     - 年份：提取文献的发表年份。
///     - 标题：为保持简洁，只取标题的前10个字符。
///     - ID：获取文献条目的全局唯一标识符（GUID）。
/// 3.  组合字符串：按照预设的格式将上述信息组合成一个完整的引用字符串。
/// 4.  复制到剪贴板：将生成的字符串复制到系统剪贴板，供用户在其他应用中粘贴使用。
/// 
/// 技术亮点：
/// - 智能处理作者列表，对于多位作者的情况，自动采用 "et al." 的学术惯例。
/// - 对标题进行截断，使得生成的引用字符串更加紧凑，适合在笔记或列表中使用。
/// - 在引用末尾附加了文献的GUID，这是实现从外部笔记（如Obsidian）精确跳转回Citavi中特定文献条目的关键。
/// 
/// 使用场景：
/// 当你在Obsidian或其他笔记软件中需要引用Citavi中的某篇文献时，只需在Citavi中选中该文献，运行此宏，然后粘贴即可。
/// 它是构建“文献-笔记”双向链接体系中，从Citavi侧生成引用链接的便捷工具。
/// </summary>


public static class CitaviMacro
{
    public static void Main() 
    {
	    // 1. 
		Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		mainForm.ActiveWorkspace = MainFormWorkspace.ReferenceEditor;
		Reference reference = mainForm.ActiveReference;
		//DebugMacro.WriteLine(reference.Title);
		//DebugMacro.WriteLine(reference.Authors.ToStringSafe());
		//DebugMacro.WriteLine(reference.YearResolved);
		//DebugMacro.WriteLine(reference.Id.ToStringSafe());
        //DebugMacro.WriteLine(reference.ShortTitle);
		//DebugMacro.WriteLine(FormatReferenceCitation(reference));
		string obsidianLink = FormatReferenceCitation(reference);
		Clipboard.SetText(obsidianLink);
		//MessageBox.Show(string.Format("链接已复制到剪贴板！\n\n{0}", obsidianLink), "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }


	/// <summary>
	/// 根据文献信息生成格式化的引用字符串。
	/// </summary>
	/// <param name="reference">要处理的文献条目对象。</param>
	/// <returns>格式化后的字符串，例如：(Author et al. Year – ShortTitle)RefID：GUID</returns>
	public static string FormatReferenceCitation(Reference reference)
	{
	    if (reference == null)
	    {
	        return "Reference object is null.";
	    }

	    // 1. 决定使用 ShortTitle 还是 Title
	    string titleToUse= reference.Title; // Title 通常不允许为null，但为安全起见可以再检查一次
	    // --- 新增：截断标题，最长10个字符 ---
	    if (titleToUse.Length > 20)
	    {
	        titleToUse = titleToUse.Substring(0, 20);
	    }
	    // 2. 获取作者信息，并处理 "et al." 的情况
	    string authorsText = reference.Authors.ToStringSafe();
	    if (reference.Authors != null && reference.Authors.Count > 1)
	    {
	        // 获取第一个作者
	        Person firstAuthor = reference.Authors.FirstOrDefault();
	        if (firstAuthor != null && !string.IsNullOrEmpty(firstAuthor.LastName))
	        {
	            authorsText = firstAuthor.LastName + " et al.";
	        }
	        else
	        {
	            // 如果第一个作者或其姓氏为空，则回退到原始的作者字符串
	            authorsText = reference.Authors.ToStringSafe();
	        }
	    }

	    // 3. 获取年份
	    string yearText = reference.YearResolved.ToString();

		string referenceGroupname ="";
		if (reference.Groups.FirstOrDefault() != null)
		{
			referenceGroupname = reference.Groups.FirstOrDefault().FullName;
		}
	    // 4. 使用 string.Format 拼接最终字符串
	    string formattedCitation = string.Format("({0} {1} – {2} {3}) RefID：{4}",
	        authorsText,
	        yearText,
	        referenceGroupname,
			titleToUse,			
	        reference.Id.ToStringSafe()
	    );

	    return formattedCitation;
	}
}