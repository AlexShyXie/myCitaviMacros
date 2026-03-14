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
using System.Reflection; // 【新增】反射需要用到
using System.Collections; // 【新增】IEnumerable需要用到

/*
 * =================================================================================================
 *
 * Citavi 知识条目链接生成宏
 *
 * =================================================================================================
 *
 * 功能概述:
 * --------
 * 本宏用于获取用户在 Citavi 中选中的知识条目，并生成包含该条目
 * 核心陈述和跳转链接的复合字符串。生成的链接采用 AutoHotkey 协议，配合外部脚本
 * 可实现从笔记软件直接跳转回 Citavi 中对应的知识条目。
 *
 * 工作流程:
 * --------
 * 1. 检测当前活动工作区，支持从文献编辑器或知识组织器
 *    中获取选中的知识条目。
 * 2. 验证用户是否精确选择了一个知识条目，确保操作的准确性。
 * 3. 提取知识条目的核心陈述作为引用文本。
 * 4. 获取当前 Citavi 项目的文件路径，用于构建跳转链接。
 * 5. 生成 AutoHotkey 协议链接，格式为：
 *    ahk://citavi/goto?type=Know&id=知识条目ID&project=项目路径
 * 6. 将核心陈述和跳转链接组合成 Obsidian Markdown 格式的字符串。
 * 7. 将生成的字符串复制到系统剪贴板，并在宏控制台输出调试信息。
 *
 * 关键技术点:
 * ----------
 * - 工作区适配：智能识别当前活动的工作区类型，支持多种获取选中项的方法。
 * - 项目路径处理：自动获取本地 SQLite 项目的完整路径，并处理路径格式转换。
 * - 链接构建：使用标准化的 AutoHotkey 协议格式，确保跳转功能的可靠性。
 * - 错误处理：包含选中项验证和项目类型检查，提供友好的错误提示。
 *
 * 使用场景:
 * --------
 * 当你在 Obsidian 或其他笔记软件中需要引用 Citavi 中的某个知识条目时，
 * 只需在 Citavi 中选中该条目，运行此宏，然后粘贴即可。生成的链接配合
 * 相应的 AutoHotkey 脚本，可以让你在笔记中点击链接直接跳转到 Citavi 中
 * 对应的知识条目，实现知识的快速定位和回溯。
 *
 * =================================================================================================
 */
public static class CitaviMacro
{
    public static void Main()
    {
        Project project = Program.ActiveProjectShell.Project;
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
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
            MessageBox.Show("请先在引文列表中精确地选择一个Knowledge。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return; // 停止执行宏
        }

        //Location location = selectedLocations[0];
        KnowledgeItem knowledgeItem = selectedKnowledgeItems[0];

        // 【修改】获取备份信息逻辑：优先PDF路径 -> 其次Reference标题 -> 最后默认值
        string backupInfo;
        string pdfPath = GetPdfPathFromKnowledgeItem(knowledgeItem); // 利用已有的方法获取PDF路径

        if (!string.IsNullOrEmpty(pdfPath))
        {
            // 优先级1：如果有PDF，获取文件名（含扩展名）
            backupInfo = System.IO.Path.GetFileName(pdfPath);
        }
        else if (knowledgeItem.Reference != null && !string.IsNullOrEmpty(knowledgeItem.Reference.Title))
        {
            // 优先级2：如果没有PDF，但有Reference，获取标题
            backupInfo = knowledgeItem.Reference.Title;
        }
        else
        {
            // 优先级3：都没有
            backupInfo = "Reference未知";
        }

        // 【修改】获取页码信息：改为从 Annotation 获取实际页码
        string pageText = "NA"; // 默认值
        SwissAcademic.Citavi.Annotation annotation = GetAnnotationFromKnowledgeItem(knowledgeItem);
        if (annotation != null)
        {
            int pageIndex = GetPageIndexSafely(annotation);
            if (pageIndex > 0)
            {
                pageText = pageIndex.ToString();
            }
        }

        // 1. 先声明 out 变量
        string projectIdentifier;
        string projectType;

        // 2. 再调用方法并传递变量
        GetProjectInfo(Program.ActiveProjectShell.Project, out projectIdentifier, out projectType);

        string ahkUrl = string.Format(
            "ahk://citavi/goto?type=Know&id={0}&project={1}&projectType={2}&Info={3}",
            knowledgeItem.Id.ToStringSafe(),
            projectIdentifier.Replace(" ", "%20"), // 项目路径空格处理
            projectType,
            backupInfo.Replace(" ", "%20") // 【关键】对新加的信息进行URL编码，防止中文乱码
        );

        // 将ahk链接包装在Obsidian的Markdown链接格式中，链接文本设为 "ahklink"
        // 【修改】将页码信息添加到链接文本中
        string obsidianLink = string.Format("[ahklink p.{0}]({1})", pageText, ahkUrl);

        string finalOutput = knowledgeItem.CoreStatement + " " + obsidianLink;
        Clipboard.SetText(finalOutput);
        DebugMacro.WriteLine(finalOutput);
    }

    /// <summary>
    /// 从知识条目获取关联的 Annotation 对象
    /// </summary>
    public static SwissAcademic.Citavi.Annotation GetAnnotationFromKnowledgeItem(KnowledgeItem knowledgeItem)
    {
        if (knowledgeItem == null) return null;

        EntityLink pdfLink = null;
        foreach (var el in knowledgeItem.Project.EntityLinks)
        {
            if (el.Source == knowledgeItem && el.Indication == "PdfKnowledgeItem")
            {
                pdfLink = el;
                break;
            }
        }

        if (pdfLink != null && pdfLink.Target is SwissAcademic.Citavi.Annotation)
        {
            return (SwissAcademic.Citavi.Annotation)pdfLink.Target;
        }
        return null;
    }

    /// <summary>
    /// 通过反射安全获取页码（来自你的第一个文件）
    /// </summary>
    public static int GetPageIndexSafely(SwissAcademic.Citavi.Annotation annotation)
    {
        if (annotation == null) return 0;

        try
        {
            // 1. 获取 Quads 属性
            var quadsProperty = annotation.GetType().GetProperty("Quads");
            if (quadsProperty == null) return 0;

            // 2. 获取属性值
            var quadsValue = quadsProperty.GetValue(annotation, null);
            if (quadsValue == null) return 0;

            // 3. 使用非泛型的 IEnumerable (System.Collections)
            var quadsEnumerable = quadsValue as IEnumerable;
            if (quadsEnumerable == null) return 0;

            // 4. 获取第一个元素
            var enumerator = quadsEnumerable.GetEnumerator();
            if (enumerator.MoveNext())
            {
                object firstQuad = enumerator.Current;
                if (firstQuad != null)
                {
                    // 5. 从 Quad 对象反射获取 PageIndex
                    var pageIndexProperty = firstQuad.GetType().GetProperty("PageIndex");
                    if (pageIndexProperty != null)
                    {
                        return (int)pageIndexProperty.GetValue(firstQuad, null);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            DebugMacro.WriteLine("获取页码失败: " + ex.Message);
        }
        return 0;
    }

    /// <summary>
    /// 获取项目的标识符（路径或名称）和类型。
    /// </summary>
    /// <param name="activeProject">当前活动的Citavi项目。</param>
    /// <param name="projectIdentifier">返回项目的标识符（本地项目为路径，服务器项目为名称）。</param>
    /// <param name="projectType">返回项目的类型（如 "DesktopSQLite" 或 "DesktopSqlServer"）。</param>
    public static void GetProjectInfo(Project activeProject, out string projectIdentifier, out string projectType)
    {
        projectIdentifier = String.Empty;
        projectType = activeProject.ProjectType.ToStringSafe();

        if (activeProject.DesktopProjectConfiguration.ProjectType == ProjectType.DesktopSQLite)
        {
            // 本地项目：标识符是文件路径
            projectIdentifier = activeProject.DesktopProjectConfiguration.SQLiteProjectInfo.FilePath;
        }
        else
        {
            // 服务器项目：标识符是项目名称
            projectIdentifier = activeProject.Name;
        }
    }

    /// <summary>
    /// 核心辅助函数：从一个知识条目对象中，解析出其关联的PDF文件的本地路径。
    /// </summary>
    /// <param name="knowledgeItem">要查询的知识条目。</param>
    /// <returns>如果找到关联的本地PDF文件，则返回其完整路径；否则返回 null。</returns>
    public static string GetPdfPathFromKnowledgeItem(KnowledgeItem knowledgeItem)
    {
        // 步骤 0: 安全检查，确保输入的知识条目不为空。
        if (knowledgeItem == null)
        {
            return null;
        }

        // 步骤 1: 通过 EntityLink 集合查找指向PDF注释的链接。
        // 链接的 Indication 必须是 "PdfKnowledgeItem"，这是Citavi标记PDF引文的标准方式。
        EntityLink pdfLink = null;
        foreach (var el in knowledgeItem.Project.EntityLinks)
        {
            if (el.Source == knowledgeItem && el.Indication == "PdfKnowledgeItem")
            {
                pdfLink = el;
                break; // 找到第一个匹配的链接后立即退出循环。
            }
        }

        // 检查是否找到了有效的链接。
        if (pdfLink == null)
        {
            return null;
        }

        // 检查链接的目标对象是否为 Annotation 类型。
        if (!(pdfLink.Target is SwissAcademic.Citavi.Annotation))
        {
            return null;
        }

        SwissAcademic.Citavi.Annotation pdfAnnotation = (SwissAcademic.Citavi.Annotation)pdfLink.Target;

        // 步骤 2: 从 Annotation 对象获取其所属的 Location 对象。
        var location = pdfAnnotation.Location;
        if (location == null)
        {
            return null;
        }

        // 步骤 3: 从 Location 对象获取其 Address 对象。
        var address = location.Address;
        if (address == null)
        {
            return null;
        }

        // 步骤 4: 调用 Address 的 Resolve() 方法，获取最终的、可访问的 Uri。
        // 这个方法能正确处理本地文件和云附件的缓存路径。
        Uri pdfUri = address.Resolve();

        // 步骤 5: 检查 Uri 是否有效且指向一个本地文件。
        if (pdfUri != null && pdfUri.IsFile)
        {
            // 获取本地文件的完整路径。
            string pdfFilePath = pdfUri.LocalPath;
            return pdfFilePath;
        }
        else
        {
            // 如果不是本地文件（例如是远程URL）或地址无效，则返回 null。
            return null;
        }
    }
}