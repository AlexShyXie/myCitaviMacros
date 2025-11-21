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
 *                               Citavi to Obsidian 链接生成宏
 *
 * =================================================================================================
 *
 * 功能概述:
 * --------
 * 本宏的核心目标是解决从 Citavi 精确跳转到 Obsidian 中对应 PDF 位置的需求。
 * 它通过读取用户在 Citavi 中选中的知识条目，找到其关联的 PDF 注释，
 * 然后提取注释的精确坐标，并将其转换为 Obsidian 支持的 `rect` 格式链接。
 * 生成的链接会自动复制到剪贴板，方便用户直接粘贴到 Obsidian 笔记中。
 *
 * 工作流程:
 * --------
 * 1.  获取用户当前在 Citavi PDF 预览中选中的知识条目。
 * 2.  通过 `EntityLink` 查找与该知识条目绑定的 `Annotation` 对象。
 *     `EntityLink` 是连接“思考”和“引用”的桥梁，其 `Indication` 属性为
 *     `PdfKnowledgeItemIndication` 时，表示这是一个标准的 PDF 引用链接。
 * 3.  从 `Annotation` 对象中提取 `Quads` 集合。`Quads` 是一系列四边形坐标，
 *     精确地描述了注释在 PDF 页面上的几何形状，可能跨越多行或不连续区域。
 * 4.  计算所有 `Quads` 的总包围盒，得到一个能覆盖整个选区的最小矩形。
 * 5.  （可选）为包围盒增加几个像素的填充，以解决视觉上高亮区域可能
 *     紧贴文字、缺乏呼吸空间的问题。
 * 6.  将计算出的坐标和页面信息，格式化为 Obsidian 的 `rect` 链接。
 *     链接格式为：`[文件名, p.页码](文件名.pdf#page=页码&rect=x1,y1,x2,y2&color=yellow)`
 * 7.  将生成的完整 Markdown 链接复制到系统剪贴板，并在 Citavi 宏控制台
 *      输出调试信息，包括链接内容和原始注释的详细信息。
 *
 * 关键技术点:
 * ----------
 * - Citavi 对象模型: 深入使用了 `KnowledgeItem`, `Annotation`, `EntityLink`,
 *   `Location` 等核心对象，理解它们之间的关系是本宏的基础。
 * - 坐标系统: Citavi 的 PDF 坐标系以左下角为原点，单位为点。Obsidian 的
 *   `rect` 参数也使用相同的单位和坐标原点，因此转换时无需进行复杂的 DPI 缩放
 *   或坐标系翻转，只需简单的数值传递和四舍五入。
 * - 链接解析: 通过 `location.Address.Resolve()` 获取 PDF 文件的本地绝对路径，
 *   并从中提取文件名用于构建链接。
 * - 错误处理: 代码中包含了多层空值检查，确保在找不到知识条目、注释或文件路径时，
 *   宏能够优雅地退出并给出提示，而不是崩溃。
 *
 * 使用场景:
 * --------
 * 当你在 Citavi 中阅读文献并做了大量高亮和笔记后，希望在 Obsidian 中撰写
 * 思考或总结时，能够一键创建一个指向 Citavi 中精确位置的链接。这个宏完美
 * 地填补了 Citavi 和 Obsidian 之间的工作流鸿沟，实现了知识的无缝引用和回溯。
 *
 * =================================================================================================
 */


public static class CitaviMacro
{
	public static void Main()
	{
		Project project = Program.ActiveProjectShell.Project;
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		string citaviFilePath = Program.ActiveProjectShell.Project.Addresses.AttachmentsFolderPath;
		DebugMacro.WriteLine(citaviFilePath);
	
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

		// 检查是否只选中了一个PDF附件位置
		//if (selectedLocations == null || selectedLocations.Count != 1)
		//{
		//	MessageBox.Show("请先在附件列表中精确地选择 **一个** PDF附件。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
		//	return; // 停止执行宏
		//}
		
		//Location location = selectedLocations[0];
		KnowledgeItem knowledgeItem = selectedKnowledgeItems[0];
		DebugMacro.WriteLine(knowledgeItem.CoreStatement);
		// 检查这个知识条目是否有任何实体链接
		// 2. 检查它是否有任何 EntityLink
		    // 在 if 语句外部声明变量，并初始化为 null
    	Annotation linkedAnnotation = null;
		if (knowledgeItem.EntityLinks != null && knowledgeItem.EntityLinks.Count() > 0)
		{
		    // 3. 在其 EntityLinks 中查找特定类型的链接
		    //    PdfKnowledgeItemIndication 是连接 Quotation 和 Annotation 的标准类型
		    var targetLink = knowledgeItem.EntityLinks
		        .Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication)
		        .FirstOrDefault();

		    // 4. 从找到的 EntityLink 中获取 Target，并转换为 Annotation
			//DebugMacro.WriteLine(targetLink != null);
			//DebugMacro.WriteLine(targetLink.Target is Annotation);
		    if (targetLink != null && targetLink.Target is Annotation)
		    {
		        linkedAnnotation = (Annotation)targetLink.Target;
		        // 现在你得到了 Annotation 对象
		    }
		}
		// 调用函数打印详情
	    if (linkedAnnotation != null)
	    {
	        // 调用新的转换函数
	        string outLink = ConvertAnnotationToObsidianLink(linkedAnnotation);
			string obsidianLink = string.Format("{0} {1} KnowID：{2}",knowledgeItem.CoreStatement,outLink,knowledgeItem.Id.ToStringSafe());
	        if (!string.IsNullOrEmpty(obsidianLink))
	        {
				PrintAnnotationDetails(linkedAnnotation);
				
	            DebugMacro.WriteLine("--- 生成的 Obsidian 链接 ---");
	            DebugMacro.WriteLine(obsidianLink);
	            DebugMacro.WriteLine("---------------------------");

	            // 可选：将链接复制到剪贴板
	            Clipboard.SetText(obsidianLink);
	            DebugMacro.WriteLine("链接已复制到剪贴板！");
				//MessageBox.Show(string.Format("链接已复制到剪贴板！：{0}",obsidianLink));
	        }
	    }
	    else
	    {
	        DebugMacro.WriteLine("未找到关联的 Annotation，无法转换。");
	    }
	

			
		
	}
	
	
	/// <summary>
	/// 将一个 Citavi Annotation 对象转换为 Obsidian 的 rect 链接格式。
	/// </summary>
	/// <param name="annotation">要转换的 Citavi Annotation 对象。</param>
	/// <returns>返回 Obsidian 链接字符串，如果转换失败则返回 null。</returns>
	public static string ConvertAnnotationToObsidianLink(Annotation annotation)
	{
		string ObsidianPdfBasePath = "11_影像学习/9_书籍PDF"; //这里设置书库路径
	    // 1. 基本校验
	    if (annotation == null) return null;

	    // 2. 获取PDF文件名
	    Location location = annotation.Location;
		
		//DebugMacro.WriteLine(location.Reference.Groups.FirstOrDefault());
	    if (location == null)
		{
		    return null;
		}
		if (location.Address == null)
		{
		    return null;
		}
		string locationGroupname ="";
		if (location.Reference.Groups.FirstOrDefault() != null)
		{
			locationGroupname = location.Reference.Groups.FirstOrDefault().FullName;
		}
		
		
	    // --- 核心修复：借鉴参考宏的逻辑，根据类型获取真实路径 ---
	    string fullFilePath = location.Address.Resolve().LocalPath;

		
	    // 最后的路径有效性检查
	    if (string.IsNullOrEmpty(fullFilePath) || !System.IO.File.Exists(fullFilePath))
	    {
	        DebugMacro.WriteLine("无法获取有效的本地文件路径或文件不存在。解析后的路径: " + fullFilePath);
	        return null;
	    }

	    // --- 调试输出，这次应该能显示正确的G盘路径了 ---
	    DebugMacro.WriteLine("最终确定的完整文件路径: " + fullFilePath);
		// 3. 从完整路径中提取文件名
	    string pdfFileName = System.IO.Path.GetFileName(fullFilePath);
	    // 4. 从完整路径中提取文件所在的父目录名
	    string parentDirectoryName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(fullFilePath));
	    // --- 核心改造点：构建Obsidian库内路径 ---
	    // 3. 将文件名和父目录名中的空格替换为 %20
	    string encodedFileName = pdfFileName.Replace(" ", "%20");
	    string encodedParentDirectoryName = parentDirectoryName.Replace(" ", "%20");

	    // 4. 组合成完整的Obsidian相对路径
	    string obsidianRelativePath = string.Format("{0}/{1}/{2}",ObsidianPdfBasePath,encodedParentDirectoryName,encodedFileName);
		//DebugMacro.WriteLine(fullFilePath);
		//DebugMacro.WriteLine(pdfFileName);
		//DebugMacro.WriteLine(parentDirectoryName);
	    // 5. 计算总包围盒
	    double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
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

	    // --- 应用像素填充 ---
	    int horizontalPadding = 2;
	    int verticalPadding = 2;
	    minX -= 2;
	    maxX += 4;//horizontalPadding;
	    minY -= 2;
	    maxY += 2;

	    // 6. 转换坐标
	    int obsX1 = (int)Math.Round(minX);
	    int obsY1 = (int)Math.Round(minY);
	    int obsX2 = (int)Math.Round(maxX);
	    int obsY2 = (int)Math.Round(maxY);

	    // 7. 获取页面索引
	    int pageIndex = annotation.Quads.First().PageIndex;

	    // 8. 生成链接
	    string linkText = string.Format("{0}, p.{1}", pdfFileName, pageIndex + 1);
	    
	    // --- 核心改造点：使用新的Obsidian路径作为链接目标 ---
	    string linkTarget = string.Format("{0}#page={1}&rect={2},{3},{4},{5}&color=yellow", 
	        obsidianRelativePath, 
	        pageIndex, 
	        obsX1, obsY1, obsX2, obsY2);

	    string finalLink = string.Format("[{0} {1}]({2})",locationGroupname,linkText, linkTarget);

	    return finalLink;
	}


	public static void PrintAnnotationDetails(Annotation annotation)
	{
	    if (annotation == null)
	    {
	       DebugMacro.WriteLine("Annotation is null.");
	        return;
	    }

	    DebugMacro.WriteLine("==================== Annotation Details ====================");
	    DebugMacro.WriteLine(string.Format("ID: {0}", annotation.Id));
	    DebugMacro.WriteLine(string.Format("Visible: {0}", annotation.Visible));

	    // 打印所属的PDF文件信息
	    Location location = annotation.Location;
	    if (location != null)
	    {
	        if (location.Address != null)
	        {
	            DebugMacro.WriteLine(string.Format("PDF File: {0}", location.Address.Resolve().LocalPath));
	        }
	        else
	        {
	            DebugMacro.WriteLine("PDF File: Not found (Address is null).");
	        }
	    }
	    else
	    {
	        DebugMacro.WriteLine("PDF File: Not found (Location is null).");
	    }
		//DebugMacro.WriteLine(annotation);
	    // 打印颜色信息
	    System.Drawing.Color color = annotation.OriginalColor;
	    DebugMacro.WriteLine(string.Format("Color (R,G,B,A): {0}, {1}, {2}, {3}", 
	        color.R, 
	        color.G, 
	        color.B, 
	        color.A));

	    // 打印位置信息
	    if (annotation.Quads != null)
	    {
	        if (annotation.Quads.Count() > 0)
	        {
	            Console.WriteLine("--- Position & Geometry ---");
	            foreach (var quad in annotation.Quads)
	            {
	                DebugMacro.WriteLine(string.Format("  -> Segment on Page {0}:", quad.PageIndex));
	                DebugMacro.WriteLine(string.Format("     Bottom-Left (X, Y): ({0:F2}, {1:F2})", quad.MinX, quad.MinY));
	                DebugMacro.WriteLine(string.Format("     Top-Right (X, Y): ({0:F2}, {1:F2})", quad.MaxX, quad.MaxY));
	            }
	        }
	        else
	        {
	            DebugMacro.WriteLine("--- Position & Geometry: Quads collection is empty. ---");
	        }
	    }
	    else
	    {
	        DebugMacro.WriteLine("--- Position & Geometry: No Quads found (Quads is null). ---");
	    }

	    DebugMacro.WriteLine("========================================================");
	}

}