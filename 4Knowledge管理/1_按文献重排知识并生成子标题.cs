
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection; // 添加反射命名空间
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;


// =================================================================================================
// 宏名称：按文献重排知识并生成子标题
// 功能描述：
// 1. 自动将当前选中的知识分类下的所有知识条目，按照文献标题进行分组排序。
// 2. 在每个文献组内部，根据知识条目在PDF中的实际页码进行二次排序，确保阅读顺序一致。
// 3. 排序完成后，会自动在每个文献的知识条目前插入一个子标题，子标题内容为自定义的引用信息。
//    格式为：作者姓氏+年份+标题首字母大写+自定义字段1+自定义字段2
// 4. 如果分类中已存在子标题，宏会提示是否覆盖。
//
// 使用场景：
// 当你从一篇文献中摘录了大量知识条目（引文、想法等）到Citavi的同一个分类下时，
// 运行此宏可以快速将它们整理成结构清晰、符合阅读顺序的笔记大纲。
// 非常适合用于文献综述或读书笔记的整理。
//
// 使用方法：
// 1. 在Citavi的知识组织器中，点击选中你想要整理的知识分类（文件夹）。
// 2. 确保该分类下包含来自不同文献的知识条目。
// 3. 运行此宏。
// 4. 宏会自动完成排序和插入子标题的操作。
//
// 注意事项：
// - 此宏会修改知识条目的顺序，建议在操作前备份项目或在一个测试分类中先试用。
// - 排序功能依赖于知识条目与PDF注释的正确链接。如果条目没有链接到PDF，排序可能不精确。
// - 子标题的生成依赖于文献信息（作者、年份、标题等），请确保这些信息填写完整。
// - 此宏不依赖外部PDF库，通过反射获取页码，兼容性更好。
//
// 作者：[Hui]
// 日期：2025年
// =================================================================================================

public static class CitaviMacro
{
    public static void Main()
    {
        Project project = Program.ActiveProjectShell.Project;
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

        var category = mainForm.GetSelectedKnowledgeOrganizerCategory();
        var knowledgeItems = category.KnowledgeItems.ToList();

        if (knowledgeItems.Count > 1)
        {
            // --- 使用精简版的排序逻辑 ---
            knowledgeItems.Sort(new SimpleKnowledgeItemComparer());

            // 将排序后的结果应用到Category中
            var firstKnowledgeItem = knowledgeItems.First();
            for (int i = 1; i < knowledgeItems.Count; i++)
            {
                category.KnowledgeItems.Move(knowledgeItems[i], firstKnowledgeItem);
                firstKnowledgeItem = knowledgeItems[i];
            }
            // --- 排序结束 ---
        }

        // 最后调用子标题生成功能
        CreateSubheadings(knowledgeItems, category, false);
    }

    // --- 扩展方法 ---
    public static List<Location> GetPDFLocations(List<KnowledgeItem> knowledgeItems)
    {
        var locations = new List<Location>();
        foreach (var knowledgeItem in knowledgeItems)
        {
            foreach (var entityLink in knowledgeItem.EntityLinks)
            {
                if (entityLink.Indication.Equals(EntityLink.PdfKnowledgeItemIndication, StringComparison.OrdinalIgnoreCase) &&
                    entityLink.Target is Annotation)
                {
                    var location = ((Annotation)entityLink.Target).Location;
                    if (location == null) continue;
                    if (location.LocationType != LocationType.ElectronicAddress) continue;
                    if (location.Address.Resolve().LocalPath.EndsWith(".pdf") == false) continue;
                    if (locations.Contains(location)) continue;
                    locations.Add(location);
                }
            }
        }
        return locations;
    }

    // --- 精简版的比较器，使用反射获取页码 ---
    public class SimpleKnowledgeItemComparer : IComparer<KnowledgeItem>
    {
        public int Compare(KnowledgeItem x, KnowledgeItem y)
        {
            // 1. 首先，按文献的短标题排序
            var xTitle = x.Reference != null ? x.Reference.ShortTitle : "";
            var yTitle = y.Reference != null ? y.Reference.ShortTitle : "";
            if (xTitle != yTitle) return xTitle.CompareTo(yTitle);

            // 2. 如果文献相同，尝试按PDF页码排序
            if (x.EntityLinks.Any() && y.EntityLinks.Any())
            {
                Annotation xAnnotation = GetAnnotationFromKnowledgeItem(x);
                Annotation yAnnotation = GetAnnotationFromKnowledgeItem(y);

                if (xAnnotation != null && yAnnotation != null)
                {
                    // 比较PDF文件路径，确保是同一个文件
                    var xAddress = xAnnotation.Location.Address.ToString();
                    var yAddress = yAnnotation.Location.Address.ToString();
                    if (xAddress != yAddress) return xAddress.CompareTo(yAddress);

                    // 比较页码（使用反射方法）
                    int xPage = GetAnnotationPage(xAnnotation);
                    int yPage = GetAnnotationPage(yAnnotation);
                    if (xPage != yPage) return xPage.CompareTo(yPage);
                }
            }

            // 3. 如果PDF信息也相同，最后按知识条目本身的页码范围排序
            try
            {
                if (x.PageRange != null && y.PageRange != null)
                {
                    var xStartPage = Decimal.Parse(x.PageRange.StartPage.OriginalString);
                    var yStartPage = Decimal.Parse(y.PageRange.StartPage.OriginalString);
                    if (xStartPage != yStartPage) return xStartPage.CompareTo(yStartPage);
                }
            }
            catch
            {
                // 如果页码解析失败，则按字符串比较
                if (x.PageRange != null && y.PageRange != null)
                {
                    var xRange = x.PageRange.OriginalString;
                    var yRange = y.PageRange.OriginalString;
                    if (xRange != yRange) return xRange.CompareTo(yRange);
                }
            }

            return 0;
        }

        // 辅助方法：从知识条目中获取注释对象
        private Annotation GetAnnotationFromKnowledgeItem(KnowledgeItem item)
        {
            if (item.EntityLinks.Any(link => link.Target is Annotation))
            {
                return item.EntityLinks.First(link => link.Target is Annotation).Target as Annotation;
            }

            // 处理评论类型的知识条目，它可能链接到另一个有注释的知识条目
            if (item.QuotationType == QuotationType.Comment && item.EntityLinks.Any())
            {
                try
                {
                    var targetItem = item.EntityLinks.First(link => link.Target is KnowledgeItem).Target as KnowledgeItem;
                    if (targetItem != null)
                    {
                        return GetAnnotationFromKnowledgeItem(targetItem);
                    }
                }
                catch { }
            }
            return null;
        }

        // 辅助方法：从注释中获取页码【使用反射，兼容 C# 4.8.1】
        private int GetAnnotationPage(Annotation annotation)
        {
            try
            {
                // 使用反射获取 Quads 属性
                var quadsProperty = annotation.GetType().GetProperty("Quads");
                if (quadsProperty != null)
                {
                    var quadsValue = quadsProperty.GetValue(annotation, null);
                    if (quadsValue != null)
                    {
                        // 使用非泛型的 IEnumerable 来遍历 (兼容 C# 4.x)
                        var quadsEnumerable = quadsValue as System.Collections.IEnumerable;
                        if (quadsEnumerable != null)
                        {
                            // 获取第一个元素
                            var enumerator = quadsEnumerable.GetEnumerator();
                            if (enumerator.MoveNext())
                            {
                                object firstQuad = enumerator.Current;
                                if (firstQuad != null)
                                {
                                    // 从 Quad 对象反射获取 PageIndex
                                    var pageIndexProperty = firstQuad.GetType().GetProperty("PageIndex");
                                    if (pageIndexProperty != null)
                                    {
                                        return (int)pageIndexProperty.GetValue(firstQuad, null);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // 如果反射失败，静默失败，返回0
            }

            return 0; // 如果没有页码信息，返回0
        }
    }

    // --- 优化后的子标题生成方法 ---
    static void CreateSubheadings(List<KnowledgeItem> knowledgeItems, Category category, bool overwriteSubheadings)
    {
        var mainForm = Program.ActiveProjectShell.PrimaryMainForm;
        var projectShell = Program.ActiveProjectShell;
        var project = projectShell.Project;

        // 【修复】先过滤掉子标题，只保留真正的知识条目
        var realKnowledgeItems = knowledgeItems
            .Where(item => item.KnowledgeItemType != KnowledgeItemType.Subheading)
            .ToList();

        var subheadings = knowledgeItems
            .Where(item => item.KnowledgeItemType == KnowledgeItemType.Subheading)
            .ToList();

        Reference currentReference = null;
        Reference previousReference = null;

        // 处理现有的子标题
        if (subheadings.Any())
        {
            if (!overwriteSubheadings)
            {
                DialogResult result = MessageBox.Show(
                    "类别 \"" + category.Name + "\" 中的知识条目列表已包含子标题。\r\n\r\n如果继续，这些子标题将被首先删除。\r\n\r\n是否继续？",
                    "Citavi",
                    MessageBoxButtons.YesNo);
                if (result == DialogResult.No) return;
            }

            foreach (var subheading in subheadings)
            {
                subheading.Categories.Remove(category);
                project.Thoughts.Remove(subheading);
            }
            projectShell.SaveAsync(mainForm);
        }

        // 【修复】重新获取 category.KnowledgeItems，因为删除子标题后列表已更新
        var currentKnowledgeItems = category.KnowledgeItems.ToList();

        // 遍历真正的知识条目（不是子标题）
        foreach (var knowledgeItem in realKnowledgeItems)
        {
            // 双重保险
            if (knowledgeItem.KnowledgeItemType == KnowledgeItemType.Subheading)
            {
                continue;
            }

            if (knowledgeItem.Reference != null)
                currentReference = knowledgeItem.Reference;

            string headingText = "无可用短标题";

            if (currentReference != null)
            {
                // 获取作者、时间、Title
                string authorName = "";
                if (currentReference.Authors != null && currentReference.Authors.Count > 0)
                {
                    authorName = currentReference.Authors[0].LastName.ToString();
                }
                
                string year = currentReference.Year;
                string IF = currentReference.CustomField1;
                string Qpart = currentReference.CustomField2;

                // 获取Title并处理
                string originalTitle = currentReference.Title ?? "";
                string[] words = originalTitle.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                
                // 所有单词，首字母大写，其余小写
                string titlePart = string.Join(" ", words.Select(word => word.First().ToString().ToUpper() + word.Substring(1).ToLower()));

                // --- 构建引用键，处理空值和下划线 ---
                List<string> parts = new List<string>();

                // 1. 添加 作者+年份 部分
                string authorYear = (authorName ?? "") + (year ?? "");
                if (!string.IsNullOrEmpty(authorYear))
                {
                    parts.Add(authorYear);
                }

                // 2. 添加 标题 部分
                if (!string.IsNullOrEmpty(titlePart))
                {
                    parts.Add(titlePart);
                }

                // 3. 添加 自定义字段1 (IF) 部分
                if (!string.IsNullOrEmpty(IF))
                {
                    parts.Add(IF);
                }

                // 4. 添加 自定义字段2 (Qpart) 部分
                if (!string.IsNullOrEmpty(Qpart))
                {
                    parts.Add(Qpart);
                }

                // 使用 "_" 连接所有非空部分
                string citationkey = string.Join("_", parts);
                
                // 如果所有字段都为空，设置默认文本
                headingText = string.IsNullOrEmpty(citationkey) ? "无可用短标题" : citationkey;
            }
            else if (knowledgeItem.QuotationType == QuotationType.None)
            {
                headingText = "想法";
            }

            // 【修复】使用重新获取的列表来查找索引
            int nextInsertionIndex = currentKnowledgeItems.IndexOf(knowledgeItem);

            if (nextInsertionIndex == -1)
            {
                // 如果找不到，说明知识条目可能已被移除，跳过
                continue;
            }

            category.KnowledgeItems.AddNextItemAtIndex = nextInsertionIndex;

            if (nextInsertionIndex == 0)
            {
                var subheading = new KnowledgeItem(project, KnowledgeItemType.Subheading)
                {
                    CoreStatement = headingText
                };
                subheading.Categories.Add(category);
                project.Thoughts.Add(subheading);
                projectShell.SaveAsync(mainForm);
                previousReference = currentReference;
                
                // 【修复】更新列表，因为添加了子标题
                currentKnowledgeItems = category.KnowledgeItems.ToList();
                continue;
            }

            if (nextInsertionIndex > 0 && (currentReference != null && currentReference != previousReference))
            {
                var subheading = new KnowledgeItem(project, KnowledgeItemType.Subheading)
                {
                    CoreStatement = headingText
                };
                subheading.Categories.Add(category);
                project.Thoughts.Add(subheading);
                projectShell.SaveAsync(mainForm);
                
                // 【修复】更新列表，因为添加了子标题
                currentKnowledgeItems = category.KnowledgeItems.ToList();
            }

            previousReference = currentReference;
        }
    }
}