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
// 宏名称：合并知识分类（已废弃）
// 功能描述：
// 根据原始注释，此宏的设计意图是：
// 1_按文献重排知识并生成子标题.cs已将原本需要1、2.1和这个2.2的宏的操作全部一次性实现了。

// 此宏的实际代码并不会“合并”或“删除”分类。它的操作是：
// 1. 获取项目中所有的知识分类。
// 2. 将所有知识分类下的所有知识条目，额外地添加到项目中的第一个知识分类里。
// =================================================================================================

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

        // 获取当前主窗体中所有被筛选出来的参考文献（注意：此变量在后续代码中并未被使用）
        List<Reference> references = mainForm.GetFilteredReferences();    

        // 获取项目中所有的知识分类，并转换为一个列表
        var categories = project.Categories.ToList();

        // 创建一个列表，只包含项目中的第一个知识分类，作为“目标分类”
        List<Category> category_first = new List<Category> { categories[0] }; // categories.Take(2).ToList(); 

        // 遍历从第二个开始的所有知识分类
        foreach (Category category in categories.Skip(1))
        {
            // 弹出消息框，显示当前正在处理的分类名称（调试用，通常应注释掉）
            // MessageBox.Show(category.Name);

            // 获取当前分类下的所有知识条目
            List<KnowledgeItem> knowledgeItems = category.KnowledgeItems.ToList();
            
            // 遍历这些知识条目
            foreach (KnowledgeItem knowledgeItem in knowledgeItems)
            {
                // 弹出消息框，显示当前正在处理的知识条目全名（调试用，通常应注释掉）
                // MessageBox.Show(knowledgeItem.FullName);

                // 【核心操作】将目标分类（第一个分类）添加到当前知识条目的分类列表中
                // 注意：这里用的是 AddRange，是“添加”，而不是“移动”或“替换”
                // 这会导致知识条目同时属于它原来的分类和第一个分类
                knowledgeItem.Categories.AddRange(category_first);

                // 【被注释掉的操作】
                // knowledgeItem.Categories.Clear(); // 如果取消注释，会先清空原分类
                // project.Categories.Remove(category); // 如果取消注释，会删除整个分类（危险操作！）
            }
        }
    }
}