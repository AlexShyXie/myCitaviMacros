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

// =================================================================================================
// 宏名称：导出知识组为Markdown文档（最终修复版，制作素材文档2.1，承接1_按文献重排知识并生成子标题.cs结束后）
// 功能描述：
// 此宏用于将选定的知识分类中的所有内容，以Markdown格式输出到Citavi的输出窗口。
// 它会为每个知识条目生成对应的Markdown文本，包括核心内容、所属子标题和关联的PDF图片链接。
//
// 使用方法：
// 1. 在Citavi的知识组织器中，点击选中你想要导出的知识分类（文件夹）。
// 2. 运行此宏。
// 3. 在Citavi的“输出”窗口（通常在底部）中查看生成的Markdown文本。
// 4. 你可以全选复制这些文本，然后粘贴到任何支持Markdown的编辑器中（如Obsidian、Typora等）。
//
// 输出格式示例：
// 知识条目的核心内容
// ### 知识条目的核心内容
// 所属的子标题
// ![](关联的PDF截图或标注图片的路径)
//
// 注意事项：
// - 此宏会跳过任何无法获取有效路径的知识条目，以确保程序稳定运行。
// - 如果没有选中知识分类就运行宏，会弹出提示并退出。
// =================================================================================================

public static class CitaviMacro
{
    public static void Main()
    {
        //Get the active project
        Project project = Program.ActiveProjectShell.Project;
        
        //Get the active ("primary") MainForm
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
        
        // 获取用户当前在知识组织器中选中的知识分类
        var category = mainForm.GetSelectedKnowledgeOrganizerCategory();

        // 【修复点1】增加检查，确保用户选中了一个分类，否则category为null会报错
        if (category == null)
        {
            MessageBox.Show("请先在知识组织器中选中一个知识分类，然后再运行此宏。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 获取该分类下的所有知识条目
        var knowledgeItems = category.KnowledgeItems.ToList();
        
        // 用于存储当前条目所属的子标题
        string relateSubheading = "";
        
        // 遍历分类中的每一个知识条目
        foreach (KnowledgeItem knowledgeitem in knowledgeItems)
        {
            // 获取知识条目关联的参考文献（可能为null）
            Reference reference = knowledgeitem.Reference;
            
            // 如果当前条目是子标题类型
            if (knowledgeitem.KnowledgeItemType == KnowledgeItemType.Subheading)
            {
                // 更新当前子标题
                relateSubheading = knowledgeitem.CoreStatement;

                DebugMacro.WriteLine("## " + knowledgeitem.CoreStatement);
            }
            else // 如果是普通的知识条目（引文、摘要、评论等）
            {
				// 素材文档1
				//DebugMacro.WriteLine("## "+knowledgeitem.CoreStatement);
				//string path = knowledgeitem.Address.DataContractFullUriString;			
		        //path = path.Replace("\\", "/");// Replace backslashes with forward slashes
		        //path = path.Replace(" ", "%20");// Replace spaces with %20
				//DebugMacro.WriteLine("![]("+path+")");
				
                // 在输出窗口中生成一个三级标题，内容为知识条目的核心陈述
				// 素材文档2.1				
				DebugMacro.WriteLine(knowledgeitem.CoreStatement);
                DebugMacro.WriteLine("### " + knowledgeitem.CoreStatement);
                
                // 输出其所属的子标题
                DebugMacro.WriteLine(relateSubheading);

                // 【修复点2】增加检查，确保知识条目有关联的地址信息
                if (knowledgeitem.Address != null)
                {
                    // 获取关联的地址（通常是PDF文件的路径或图片的URI）
                    string path = knowledgeitem.Address.DataContractFullUriString;
                    
                    // 【最终修复点】增加对path本身的null检查，处理Address存在但路径为空的情况
                    if (!string.IsNullOrEmpty(path))
                    {
                        // 对路径进行格式化，使其符合Markdown图片链接的语法要求
                        // 将反斜杠替换为正斜杠
                        path = path.Replace("\\", "/");
                        // 将空格替换为URL编码
                        path = path.Replace(" ", "%20");
                        
                        // 在输出窗口中生成Markdown格式的图片链接
                        DebugMacro.WriteLine("![](" + path + ")");
                    }
                    // 如果path为null或空，则什么都不做，静默跳过
                }
                // 如果Address为null，也静默跳过
            }
        }

        // 完成后，可以给用户一个提示
        MessageBox.Show("Markdown文本已生成完毕，请在“输出”窗口中查看。", "导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}