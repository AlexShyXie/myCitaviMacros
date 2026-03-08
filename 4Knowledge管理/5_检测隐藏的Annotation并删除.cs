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

public static class CitaviMacro
{
    public static void Main()
    {
        Project project = Program.ActiveProjectShell.Project;
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

        // 获取所有知识条目
        ProjectAllLocationsCollection allLocations = project.AllLocations;
        
		// 统计变量
	        int totalChecked = 0;
	        int processedCount = 0;
	        int totalHiddenFound = 0;
	        List<string> processedItems = new List<string>();

	        // 遍历所有location
	        foreach (Location item in allLocations)
	        {
				string path = item.Address.Resolve().LocalPath;
				//DebugMacro.WriteLine(location.Address.Resolve().LocalPath);
				if (!path.EndsWith(".pdf")) continue;
	            totalChecked++;
	            int hiddenCount = 0;
	            
	            try
	            {
	                // 检查是否有隐藏注释
	                if (item.Annotations != null && item.Annotations.Count > 0)
	                {
	                    // 统计隐藏注释数量
	                    foreach (var annotation in item.Annotations)
	                    {
	                        if ((annotation.Visible == false) && (GetKnowledgeFromAnnotation(annotation) == null))
	                        {
	                            hiddenCount++;
	                        }
	                    }
	                    
	                    // 如果发现隐藏注释，则执行删除
	                    if (hiddenCount > 0)
	                    {
	                        totalHiddenFound += hiddenCount;
	                        item.RemoveInvisibleAnnotations();
	                        processedCount++;
	                        
	                        // 记录处理过的条目
	                        string result = string.Format("已处理: {0} (发现 {1} 个隐藏注释)", 
	                            item.FullName ?? "未命名条目", hiddenCount);
	                        processedItems.Add(result);
	                        
	                        DebugMacro.WriteLine(result);
	                    }
	                }
	            }
	            catch (Exception ex)
	            {
	                DebugMacro.WriteLine("处理 {0} 时出错: {1}", item.FullName, ex.Message);
	            }
	        }

	        // 生成报告
	        string report = string.Format(
	            "隐藏注释清理完成！\n\n" +
	            "总共检查: {0} 个PDF条目\n" +
	            "发现隐藏注释: {1} 个\n" +
	            "处理了: {2} 个条目\n\n",
	            totalChecked, totalHiddenFound, processedCount);

	        if (processedItems.Count > 0)
	        {
	            report += "处理详情:\n";
	            int maxDisplay = Math.Min(processedItems.Count, 15);
	            for (int i = 0; i < maxDisplay; i++)
	            {
	                report += "• " + processedItems[i] + "\n";
	            }
	            
	            if (processedItems.Count > maxDisplay)
	            {
	                report += string.Format("... 还有 {0} 个条目已处理\n", processedItems.Count - maxDisplay);
	            }
	        }
	        else
	        {
	            report += "没有发现需要处理的隐藏注释。";
	        }
	        
	        // 显示报告
	        MessageBox.Show(report, "清理报告", MessageBoxButtons.OK, MessageBoxIcon.Information);
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