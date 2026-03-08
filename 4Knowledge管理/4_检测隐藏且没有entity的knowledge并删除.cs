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
	
// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
	    Project project = Program.ActiveProjectShell.Project;
	    MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

	    // 1. 获取所有知识条目
	    ProjectAllKnowledgeItemsCollection allKnowledgeItems = project.AllKnowledgeItems;
	    
	    // 2. 创建一个列表，用于存放所有需要删除的知识条目
	    List<KnowledgeItem> itemsToDelete = new List<KnowledgeItem>();

	    // 3. 第一遍遍历：只找出需要删除的条目，并加入待删除列表
	    foreach (KnowledgeItem item in allKnowledgeItems)
	    {
	        // 使用你指定的判断条件
	        if ((item.QuotationType == QuotationType.Highlight) && (item.EntityLinks.Count() == 0))
	        {
	            // 先不删除，只是把它加入“待删除”列表
	            itemsToDelete.Add(item);
	        }
	    }

	    // 4. 第二遍处理：遍历“待删除列表”，执行真正的删除操作
	    int deletedCount = 0;
	    foreach (KnowledgeItem itemToRemove in itemsToDelete)
	    {
	        // 尝试从正确的集合中删除
	        if (itemToRemove.Reference != null)
	        {
	            itemToRemove.Reference.Quotations.Remove(itemToRemove);
	            deletedCount++;
	        }
	    }

	    // 5. 循环结束后，统一向用户报告结果
	    if (deletedCount > 0)
	    {
	        string message = string.Format("操作完成。共删除了 {0} 个没有关联实体的高亮知识条目。", deletedCount);
	        MessageBox.Show(message, "清理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
	    }
	    else
	    {
	        string message = string.Format("没有找到需要删除的（隐藏且没有实体的高亮）知识条目。");
	        MessageBox.Show(message, "无需清理", MessageBoxButtons.OK, MessageBoxIcon.Information);
	    }
	}


}