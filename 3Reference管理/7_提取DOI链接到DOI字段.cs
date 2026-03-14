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

/*
 * 宏名称：从PDF路径中提取并填充DOI字段
 * 
 * 功能描述：
 * 此宏用于处理在 Citavi 中选定的参考文献条目。它会遍历每个条目关联的电子文件路径（URL），
 * 查找其中是否包含 "https://doi.org/" 格式的 DOI 链接。如果找到，它会提取出纯 DOI 号，
 * 并自动填充到该条目的“DOI”字段中。
 * 
 * 使用场景：
 * 当你从外部（如 BookxNote Pro 或其他PDF阅读器）将文献导入 Citavi 时，有时 PDF 文件的路径
 * 或链接可能直接包含了 DOI 信息，但 DOI 字段却是空的。运行此宏可以一键批量补全这些信息，
 * 省去手动复制粘贴的麻烦。
 * 
 * 操作步骤：
 * 1. 在 Citavi 中，选中一个或多个需要处理 DOI 的文献条目。
 * 2. 运行此宏。
 * 3. 宏会自动检查并更新这些条目的 DOI 字段。
 * 
 * 注意事项：
 * - 此宏仅处理选中的条目，不会影响项目中的其他文献。
 * - 它只查找路径中 "https://doi.org/" 格式的链接。
 * - 建议在运行前备份项目，以防万一。
 */


public static class CitaviMacro
{
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		//if this macro should ALWAYS affect all titles in active project, choose:
		//ProjectReferenceCollection references = project.References;		

		//if this macro should affect just filtered rows in the active MainForm, choose:
		List<Reference> references = mainForm.GetSelectedReferences();	

		foreach (Reference reference in references)
		{
			// your code
			DebugMacro.WriteLine(reference.Doi);
			List<Location> refLocation= reference.Locations.ToList();
            foreach (Location location in refLocation)
            {
	            if (reference.Locations == null) continue;
				if (string.IsNullOrEmpty(location.Address.ToString())) continue;
				if (location.LocationType != LocationType.ElectronicAddress) continue;
				string filePath = location.Address.ToString();
				if (filePath.Contains("https://doi.org/"))
				{
					DebugMacro.WriteLine(filePath);
					DebugMacro.WriteLine(location.LocationType.ToString());
					reference.Doi = filePath.Replace("https://doi.org/", "");
				}
				DebugMacro.WriteLine(reference.Doi);

			}
		}
	}
}