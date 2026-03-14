using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;
using System.Linq;
using System.Text.RegularExpressions;


using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

// 功能：
// 这个宏的作用是筛选出所有没有关联到任何条目（如参考文献、作者、关键词等）的知识项。
// 简单来说，它会帮你找出那些"孤立"的笔记、引文或思考，这些内容没有和你的文献库建立联系。

// 具体逻辑：

// 遍历项目中的所有知识项。
// 跳过高亮和纯文本类型（因为它们通常不需要链接）。
// 检查每个知识项的EntityLinks属性，如果为空或数量为0，就把它加入一个列表。
// 最后，在Citavi的知识组织器中应用一个过滤器，只显示这些"无链接"的知识项。
// 使用注意事项：

// 运行前确保你有一个Citavi项目是打开的。
// 运行后，它会弹窗告诉你找到了多少个这样的知识项，并自动切换到过滤视图。
// 这个功能特别适合用来清理项目，找到那些你可能忘记分类的笔记，然后手动为它们链接到正确的文献或概念上。

public static class CitaviMacro
{
	public static void Main()
	{
		//****************************************************************************************************************
		// Select Knowledge Items without Entity Links (Annotations) 
		// 1.0 - 2020-07-28
		//
		//
		// EDIT HERE
		// Variables to be changed by user
			
				
		// DO NOT EDIT BELOW THIS LINE
		// ****************************************************************************************************************

		if (Program.ProjectShells.Count == 0) return;		//no project open
				
		//reference to active Project
		Project activeProject = Program.ActiveProjectShell.Project;
		
		List<KnowledgeItem> allKis = activeProject.AllKnowledgeItems.ToList();
		
		List<KnowledgeItem> noEntityLink = new List<KnowledgeItem>();
		
		try
		{
			foreach (KnowledgeItem ki in allKis)
			{
				if (ki.QuotationType == QuotationType.None || ki.QuotationType == QuotationType.Highlight) continue;
				if (ki.EntityLinks == null || ki.EntityLinks.Count() == 0) noEntityLink.Add(ki);			

			}
			

			if (noEntityLink.Count > 0)	
			{
				noEntityLink = noEntityLink.Distinct().ToList();
				KnowledgeItemFilter kiFilter = new KnowledgeItemFilter(noEntityLink, "Knowledge items without entity links", false);
				List<KnowledgeItemFilter> kiFilters = new List<KnowledgeItemFilter>();
				kiFilters.Add(kiFilter);
				Program.ActiveProjectShell.PrimaryMainForm.KnowledgeOrganizerFilterSet.Filters.ReplaceBy(kiFilters);
			}
		}
		catch (Exception e)
		{
			DebugMacro.WriteLine("An error occurred: " + e.Message);
		}		
		finally
		{
			if (noEntityLink.Count == 0)
			{
				MessageBox.Show("No such KI found.", "Macro", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			else
			{
				MessageBox.Show(string.Format("{0} reference(s) in new selection", noEntityLink.Count), "Macro", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}
	}
}