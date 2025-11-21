using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
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
        // 获取当前项目和主窗口的引用
		Project project = Program.ActiveProjectShell.Project;
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;


        // 2. 获取所有选中的条目
        //List<Reference> selectedReferences = mainForm.GetSelectedReferences();
		
		List<KnowledgeItem> selectedKnowledgeItems = null;

		// KnowledgeItem根据当前活动的工作区类型选择合适的方法
		if (mainForm.ActiveWorkspace.ToString() == "ReferenceEditor")
		{
			selectedKnowledgeItems = mainForm.GetSelectedQuotations();
		}
		else if (mainForm.ActiveWorkspace.ToString() == "KnowledgeOrganizer")
		{
			selectedKnowledgeItems = mainForm.GetSelectedKnowledgeItems();
		}
		
        //List<Location> selectedLocations = mainForm.GetSelectedElectronicLocations();

        // 3. 初始化一个空字符串来存放结果
        string output = "";
        bool foundAnyItem = false;


        // 4. 处理选中的知识条目
        if (selectedKnowledgeItems != null && selectedKnowledgeItems.Count > 0)
        {
            foundAnyItem = true;
            foreach (KnowledgeItem knowledgeItem in selectedKnowledgeItems)
            {
                output += string.Format("{0} KnowID：{1}\n",knowledgeItem.CoreStatement,knowledgeItem.Id.ToStringSafe());
            }
        }


        // 7. 根据是否找到条目来显示结果
        if (foundAnyItem)
        {
            // 移除末尾可能多余的换行符
            string finalOutput = output.TrimEnd('\n'); 
            DebugMacro.WriteLine("--- 宏执行结果 选中的知识条目 ---\n");
            DebugMacro.WriteLine(finalOutput);

            Clipboard.SetText(finalOutput);
            //MessageBox.Show(string.Format("已找到并复制选中条目的ID！\n\n{0}", finalOutput), "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("未选中任何参考文献、知识条目或附件位置。\n\n请在Citavi中选中至少一个条目后重试。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}