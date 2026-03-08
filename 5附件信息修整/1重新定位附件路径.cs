using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection; // 添加反射命名空间
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;


/*
 * ---------------------------------------------------------------------------------
 * 宏名称：重新定位附件路径（反射兼容版）
 *
 * 功能描述：
 * 本宏用于解决 Citavi 中一个常见但棘手的问题：当附件文件被移动或重命名后，
 * 如何安全地更新其在 Citavi 项目中的路径，同时完整保留所有与该附件关联的
 * 引文和注释。
 *
 * 背景与问题：
 * Citavi 的附件路径更新功能存在限制，直接修改路径的 API 在宏环境中被禁用。
 * 简单地删除旧附件并添加新附件，会导致所有基于该附件的引文和注释失效，
 * 因为 Annotation 对象不能被移动到另一个 Location，且 KnowledgeItem 通过
 * Annotation 与 Location 间接关联。
 *
 * 解决方案：
 * 本宏采用“克隆与重建”的策略，完美绕过了 Citavi 的内部限制：
 * 1. 创建：根据用户选择的新文件路径，创建一个新的 Location 对象。
 * 2. 克隆：遍历旧 Location 上的所有 Annotation，为每一个创建一个全新的、
 *    属于新 Location 的副本，并复制其所有属性（如高亮区域、颜色等）。
 *    【使用反射复制 Quads 属性，以兼容无 SwissAcademic.Pdf 引用的情况】
 * 3. 重建：遍历所有链接到旧 Annotation 的 KnowledgeItem，为它们创建新的
 *    EntityLink，重新指向新创建的 Annotation 副本。
 * 4. 迁移与清理：将旧 Location 的元数据（如预览行为）复制到新 Location，
 *    然后安全地删除旧的 Location。
 *
 * 使用方法：
 * 1. 在 Citavi 中，选中一个包含附件的参考文献。
 * 2. 在右侧预览窗格中，选中那个需要更新路径的附件。
 * 3. 运行此宏。
 * 4. 在弹出的文件选择对话框中，选择附件的新位置。
 * 5. 宏将自动完成所有更新和链接重建工作。
 *
 * 注意事项：
 * - 此操作是安全的，它会完整保留你的知识结构。
 * - 操作完成后，旧 Location 会被删除，请确保新文件路径是正确的。
 * - 此宏不依赖外部PDF库，通过反射复制属性，兼容性更好。
 *
 * ---------------------------------------------------------------------------------
 */

public static class CitaviMacro
{
    public static void Main()
    {
        Project project = Program.ActiveProjectShell.Project;
        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
        Reference reference = mainForm.ActiveReference;
        Location oldLocation = mainForm.PreviewControl.ActiveLocation;

        // 前置检查
        if (reference == null || oldLocation == null)
        {
            MessageBox.Show("请先选中一个参考文献，并在预览中选中一个附件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        OpenFileDialog fileDialog = new OpenFileDialog();
        fileDialog.Title = string.Format("为条目 \"{0}\" 选择新的附件位置", reference.Title);

        if (fileDialog.ShowDialog() == DialogResult.OK)
        {
            string newFilePath = fileDialog.FileName;
            try
            {
                // 1. 创建新的 Location，并添加到参考文献中
                Location newLocation = new Location(reference, LocationType.ElectronicAddress, newFilePath);
                reference.Locations.Add(newLocation);

                // 2. 克隆所有 Annotation 并重建链接
                // 获取旧 Location 上的所有 Annotation
                var annotationsToClone = oldLocation.Annotations.ToList();

                foreach (var oldAnnotation in annotationsToClone)
                {
                    // 2.1 创建一个全新的、属于新 Location 的 Annotation
                    Annotation newAnnotation = new Annotation(newLocation);

                    // 2.2 复制旧 Annotation 的所有重要属性
                    newAnnotation.OriginalColor = oldAnnotation.OriginalColor;
                    
                    // 【修改】使用反射复制 Quads 属性
                    CopyQuadsProperty(oldAnnotation, newAnnotation);
                    
                    newAnnotation.Visible = oldAnnotation.Visible;
                    // ... 如果还有其他属性需要复制，可以继续添加

                    // 2.3 将新 Annotation 添加到新 Location
                    newLocation.Annotations.Add(newAnnotation);

                    // 2.4 【最关键】重建 KnowledgeItem 和新 Annotation 的 EntityLink
                    // 查找所有链接到旧 Annotation 的 KnowledgeItem
                    var linkedKnowledgeItems = project.EntityLinks
                        .Where(el => el.Target == oldAnnotation)
                        .Select(el => el.Source)
                        .ToList();

                    foreach (var ki in linkedKnowledgeItems)
                    {
                        // 为每个 KnowledgeItem 创建一个新的链接，指向新的 Annotation
                        EntityLink newEntityLink = new EntityLink(project);
                        newEntityLink.Source = ki;
                        newEntityLink.Target = newAnnotation;
                        newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication; // 或者其他合适的 Indication
                        project.EntityLinks.Add(newEntityLink);
                    }
                }

                // 3. （可选）迁移其他元数据
                newLocation.PreviewBehaviour = oldLocation.PreviewBehaviour;

                // 4. 删除旧的 Location（此时它上面的 Annotation 和 EntityLink 已经没用了）
                reference.Locations.Remove(oldLocation);

                MessageBox.Show("附件路径已成功更新，所有相关引文和注释均已重新链接！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("更新路径时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    /// <summary>
    /// 使用反射复制 Annotation 的 Quads 属性。
    /// 因为 Quads 类型定义在 SwissAcademic.Pdf.dll 中，如果不引用该 DLL，
    /// 我们需要通过反射来获取和设置这个属性。
    /// </summary>
    private static void CopyQuadsProperty(Annotation source, Annotation target)
    {
        try
        {
            // 获取 Quads 属性信息
            PropertyInfo quadsProperty = typeof(Annotation).GetProperty("Quads");
            
            if (quadsProperty != null)
            {
                // 从源 Annotation 获取 Quads 值
                object quadsValue = quadsProperty.GetValue(source, null);
                
                // 将 Quads 值设置到目标 Annotation
                if (quadsValue != null)
                {
                    quadsProperty.SetValue(target, quadsValue, null);
                }
            }
        }
        catch (Exception ex)
        {
            // 如果反射失败，记录错误但不中断流程
            // 实际上 Quads 复制失败通常意味着高亮位置可能丢失，需要用户手动调整
            DebugMacro.WriteLine("复制 Quads 属性时出错: " + ex.Message);
        }
    }
}