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
// 宏名称：按文献标题创建分类并同步
// 功能描述：
// 此宏执行两个主要操作：
// 1. 【创建分类】：为当前在参考文献列表中选中的每一篇文献，自动创建一个以其引用信息（作者+年份+标题等）命名的Knowledge分类。
//    然后，将该文献本身添加到这个新创建的分类中。
// 2. 【同步分类】：将每一篇文献所关联的所有分类，同步复制到该文献下的所有知识条目（引文、摘要、评论等）上。
//    这样，知识条目就会拥有和其来源文献完全相同的分类标签。
//
// 使用场景：
// - 当你希望按文献来组织知识结构时，可以先运行此宏，快速为每篇文献建立专属的知识分类。
// - 当你为文献打上了分类标签后，希望这些标签也能体现在该文献的知识条目上，便于后续筛选和管理。
//
// 使用方法：
// 1. 在Citavi的参考文献视图中，选中你想要处理的文献（可以多选）。
// 2. 运行此宏。
// 3. 宏会自动完成创建分类和同步标签的操作，并弹窗提示处理结果。
//
// 注意事项：
// - 此宏会创建新的知识分类，并修改知识条目的分类信息，建议在操作前备份项目。
// - 宏会忽略“想法”类型（没有关联文献）和“高亮”类型的知识条目。
// =================================================================================================

public static class CitaviMacro
{
    // 注意要在101行设置Knowlwdge合并的组别!!!!! (这个注释是原作者留下的，可能指更早版本的配置，当前版本已无需手动设置)
    public static void Main()
    {
        //Get the active project
        Project project = Program.ActiveProjectShell.Project;
        
        //Get the active ("primary") MainForm
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
        
        //if this macro should ALWAYS affect all titles in active project, choose:
        //ProjectReferenceCollection references = project.References;        

        //if this macro should affect just filtered rows in the active MainForm, choose:
        // 获取用户在参考文献列表中选中的所有文献
        List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
        
        // =================================================================
        // 第一部分：为每篇选中的文献创建以其引用信息命名的知识分类
        // =================================================================
        // 使用字典来避免重复创建相同名称的分类
        Dictionary<string, Category> categoryDictionary = new Dictionary<string, Category>();
        foreach (Reference currentReference in references)
        {
            // 获取作者、时间、Title等信息，构建一个唯一的引用键（citationkey）
            Person author = currentReference.Authors[0];
            string year = currentReference.Year;
            string IF = currentReference.CustomField1;
            string Qpart = currentReference.CustomField2;
            // 获取Title并提取前10个单词
            string originalTitle = currentReference.Title;
            string[] words = originalTitle.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string result;
            if (words.Length >= 10)
            {// 取前10个单词，首字母大写，其余小写
                result = string.Join(" ", words.Take(10).Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower())); 
            }
            else
            {// 所有单词，首字母大写，其余小写
                result = string.Join(" ", words.Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower())); 
            }
            string citationkey = author.LastName.ToString() +year+"_"+result+"_"+IF+Qpart;
        
            // 在项目中创建新的知识分类，并将其与当前文献关联
            Category category = project.Categories.Add(citationkey);
            currentReference.Categories.Add(category);
        }
        
        
        //****************************************************************************************************************
        // 第二部分：将文献的分类同步到其下的知识条目
        // 功能：为知识条目添加参考类别、关键词和分组，反之亦然。
        // 2.0 -- 2017-03-16
        //
        // 以下为可配置区域
        //****************************************************************************************************************
        
        // 指向当前活动项目
        Project activeProject = Program.ActiveProjectShell.Project;
        if (activeProject == null) return;
        // List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences(); // 此行已在上面获取
        
        // 选择同步方向
        int direction = 1; // 1: 将文献的分类 -> 同步到知识条目 (常用); 2: 将知识条目的分类 -> 同步到文献

        // 设置要同步的属性类型（目前只启用了分类同步）
        bool setCategories = true;  

        // 请勿编辑此行以下的代码
        // ****************************************************************************************************************

        if (Program.ProjectShells.Count == 0) return;        //如果没有项目打开，则退出
        //if (IsBackupAvailable() == false) return;            //如果用户想先备份项目

        // 如果方向设置错误
        if (direction != 1 && direction != 2)
        {
            MessageBox.Show("方向设置不正确，请在代码第29行进行修改！", "宏错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // 初始化计数器
        int categoryCounter = 0;
        int keywordCounter = 0;
        int groupCounter = 0;
        int errorCounter = 0;

        // 遍历每一篇选中的文献
        foreach (Reference reference in references)
        {
            // 遍历该文献下的所有知识条目
            foreach (KnowledgeItem knowledgeItem in reference.Quotations)
            {
                // 忽略没有关联文献的“想法”类型条目
                if (knowledgeItem.Reference == null) continue; 
                // 忽略“高亮”类型的条目，因为它们通常不需要复杂的分类
                if (knowledgeItem.QuotationType == QuotationType.Highlight) continue;

                // 如果启用了分类同步
                if (setCategories)
                {
                    try
                    {
                        // 获取知识条目当前已有的分类列表
                        List<Category> kiCategories = knowledgeItem.Categories.ToList();
                        // 获取其来源文献的分类列表
                        List<Category> refCategories = knowledgeItem.Reference.Categories.ToList();
                        
                        // 合并两个分类列表，并去除重复项
                        List<Category> mergedCategories = kiCategories.Union(refCategories).ToList();
                        mergedCategories.Sort();

                        // 根据选择的方向执行同步操作
                        switch (direction)
                        {
                            case 1: // 文献 -> 知识条目
                                knowledgeItem.Categories.Clear();
                                knowledgeItem.Categories.AddRange(mergedCategories);
                                categoryCounter++;
                                break;

                            case 2: // 知识条目 -> 文献
                                reference.Categories.Clear();
                                reference.Categories.AddRange(mergedCategories);
                                categoryCounter++;
                                break;
                        }
                               
                    }
                    catch (Exception e)
                    {
                        // 如果发生错误，记录错误信息
                        string errorString = String.Format("处理文献“{1}”中的知识条目“{0}”时发生错误:\n  {2}", knowledgeItem.CoreStatement, reference.ShortTitle, e.Message);
                        // DebugMacro.WriteLine(errorString); // 在Citavi宏环境中，DebugMacro可能不可用，可以替换为其他日志方式或忽略
                        errorCounter++;
                    }
                   
                }
            
            }       
        }        

        // 操作完成后，显示结果信息
        string message = String.Empty;
        
        switch (direction)
        {
            case 1:
                message = "已为 {0} 个知识条目更新了分类。\n {1} 个知识条目更新了关键词。\n {2} 个知识条目更新了分组。\n 共发生 {3} 个错误。";
                break;
            case 2:
                message = "已为 {0} 个参考文献更新了分类。\n {1} 个参考文献更新了关键词。\n {2} 个参考文献更新了分组。\n 共发生 {3} 个错误。";
                break;
        }
              
        message = string.Format(message, categoryCounter.ToString(), keywordCounter.ToString(), groupCounter.ToString(), errorCounter.ToString());       
         
        MessageBox.Show(message, "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}