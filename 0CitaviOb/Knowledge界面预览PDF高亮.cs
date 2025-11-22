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
using System.IO;
// =================================================================================================
// Citavi 宏：通过剪贴板中的ID定位知识条目并在PDF中跳转
// =================================================================================================
// 
// 功能概述：
// 此宏旨在实现一个高效的工作流：用户在Word或其他外部工具中复制一个或多个Citavi知识条目的ID，
// 然后运行此宏。宏会自动：
// 1. 从剪贴板读取ID。
// 2. 在当前Citavi项目中查找对应的知识条目。
// 3. 将Citavi界面切换到第一个找到的知识条目。
// 4. 如果该知识条目关联了PDF文件，则在PDF预览中精确跳转到该条目对应的高亮或注释位置。
//
// 使用方法：
// 1. 在Citavi中，打开“工具” -> “宏” -> “编辑宏”，将此代码粘贴进去并保存。
// 2. 将一个或多个知识条目的ID（格式为GUID，例如 `09e2e088-f2e6-45d5-a03c-528d82ad64e9`）复制到剪贴板，每个ID占一行。
// 3. 在Citavi中运行此宏。
//
// 注意事项：
// - 宏的实现依赖于Citavi的对象模型，该模型在未来版本中可能会发生变化。
// - 如果知识条目没有关联PDF，或者关联的PDF文件无法访问，宏将弹出提示。
// - 此宏优先处理剪贴板中的第一个有效ID。
//
// =================================================================================================

public static class CitaviMacro
{
    /// <summary>
    /// 宏的主入口点，定义为异步方法以支持PDF跳转等异步操作。
    /// </summary>
    public static async Task Main() 
    {
        // --- 1. 初始化项目与主窗口引用 ---
        // 获取当前活动的Citavi项目和主窗口对象，以便进行后续操作。
        Project project = Program.ActiveProjectShell.Project;		
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
        // --- 2. 从剪贴板读取ID并查找知识条目 ---
        // 获取项目中所有的知识条目，用于后续查找。
        List<KnowledgeItem> foundKnowledgeItems = mainForm.GetSelectedKnowledgeItems();


		//MessageBox.Show(foundKnowledgeItems.Count.ToString());
        // --- 3. 调试代码块 ---
        // 此区域用于开发时的调试输出，在正式使用时可以注释掉。
         foreach (KnowledgeItem knowledge in foundKnowledgeItems)
        {
			DebugMacro.WriteLine(string.Format("找到知识条目名字: {0}", knowledge.FullName));
			DebugMacro.WriteLine(string.Format("核心思想: {0}", knowledge.CoreStatement));
			DebugMacro.WriteLine(string.Format("ID: {0}", knowledge.Id.ToString()));
		}

        // --- 4. Citavi界面操作与跳转 ---
        // 检查是否成功找到了至少一个知识条目。
        if(foundKnowledgeItems.Count > 0)
        {
			//return;
            // 4.1. (可选) 在知识组织器中应用筛选
            // 为了让用户看到所有匹配的条目，可以在知识组织器中创建一个临时筛选器。
            // 这一步不是跳转所必需的，但能提供更好的上下文。
            //mainForm.ActiveWorkspace = MainFormWorkspace.KnowledgeOrganizer;
            //KnowledgeItemFilter filter = new KnowledgeItemFilter(foundKnowledgeItems, "Knowledge Items Selected in Word", false);
            //Program.ActiveProjectShell.PrimaryMainForm.KnowledgeOrganizerFilterSet.Filters.ReplaceBy(new List<KnowledgeItemFilter> { filter });
        }
        else
        {
            // 如果剪贴板中没有找到任何有效的ID，则直接退出宏。
            MessageBox.Show("剪贴板中未找到有效的知识条目ID。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
		KnowledgeItem targetKnowledgeItem = foundKnowledgeItems[0];
        mainForm.ActiveKnowledgeItem = targetKnowledgeItem;


		// 3. 核心：通过遍历控件树来找到 QuotationSmartRepeater
		SmartRepeater<KnowledgeItem> KnowledgeItemSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("SmartRepeater", true).FirstOrDefault() as SmartRepeater<KnowledgeItem>;
		QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater =
		Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("knowledgeItemPreviewSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;


		// 3. 核心：通过遍历控件树来找到 QuotationSmartRepeater，旧版
        //Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
        //SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

		if (quotationSmartRepeaterAsQuotationSmartRepeater != null)
		{
		    // 调用方法进行高亮
		    quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate(targetKnowledgeItem, true);
		}
		
        // --- 5. 执行PDF跳转 ---
        // 检查目标知识条目是否确实关联了一个有效的地址（通常是PDF文件）。
        if (targetKnowledgeItem.Address != null)
        {
            // 调用核心方法，在PDF预览中异步跳转到知识条目对应的位置。
            await Program.ActiveProjectShell.PrimaryMainForm.PreviewControl.ShowPdfLinkAsync(mainForm, targetKnowledgeItem);
			Program.ActiveProjectShell.ShowMainForm();
        }
        else
        {
            // 如果知识条目没有关联地址（例如，它是一个纯文本的摘要），则提示用户。
            MessageBox.Show("知识条目 '{targetKnowledgeItem.CoreStatement}' 没有关联的PDF文件。", "无法跳转", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}