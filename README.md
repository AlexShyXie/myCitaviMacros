# myCitaviMacros
This is the repo of my custom Macros of citavi 6.8

2025-11-08更新，更完善，需要的同学用GPT让ai总结一下就好了，非常简单。


## **Citavi 宏功能分类说明**

### **`0CitaviOb`：Citavi 与 Obsidian 深度联动**
这是整个宏集合的精华部分，专注于打通 Citavi 和 Obsidian 之间的数据壁垒，构建双向知识流。
*   **数据获取与预览系列**：
    *   `0从剪切版获取AnnotID并在预览中显示.cs`
    *   `1从剪切版获取Knowledge并在预览中显示.cs`
    *   `2从剪切版获取Reference并在预览中显示.cs`
    *   `3从剪切版获取PDF路径并在预览中显示.cs`
    *   **功能**：通过复制 ID 到剪贴板，快速在预览窗口中获取对应的标注、知识条目、文献信息或 PDF 路径，是联动的基础。
*   **信息提取与转换系列**：
    *   `4获取选中的reference文献信息.cs`
    *   `导出文献信息为YAML格式.cs`
    *   **功能**：直接提取当前选中文献的详细信息，并可导出为 YAML 格式，方便在 Obsidian 等工具中作为元数据使用。
*   **坐标转换核心系列**：
    *   `5从knowledge获取annotation并将quad转换成pdf++rect坐标.cs`
    *   `6从annotation获取quad和knowledge信息并转换成pdf++rect坐标.cs`
    *   **功能**：解决不同软件间 PDF 标注坐标不兼容的关键技术。将 Citavi 的 `quad` 坐标转换为 PDF++ 等工具能识别的 `rect` 坐标，实现高亮区域的精准同步。
*   **辅助与调试**：
    *   `7单纯选中knowledge获取CoreState和ID.cs`：用于调试，获取知识条目的核心状态和 ID。
    *   `参考：高亮Knowledge条目.cs`：一个参考实现，用于高亮显示知识条目。
    *   `失败CopyTextOfSelectedAnnotation.cs`：一个失败的尝试记录，可能用于调试或作为反面教材。
    *   `获取选中的annotation-knowledge信息-...版.cs`：多个版本的同功能宏，可能针对不同的预览场景（全屏、右侧面板等）进行了优化。


### **`1影响因子及翻译`：文献信息增强**
利用外部服务和 AI，批量丰富和翻译文献元数据。
*   **`easyScholarIF影响因子.cs`**：调用 easyScholar 等服务，批量获取文献的影响因子。
*   **AI 翻译系列**：
    *   `2ollamaWin10_AbstractAndTitleOfSelected.cs`
    *   `3智谱4flash_AbstractAndTitleOfSelected.cs`
    *   `ChatGLM3_AbstractAndTitleOfSelected.cs`
    *   `ChatGLM3_TitleOfSelected.cs`
    *   **功能**：集成不同的 AI 模型（本地 Ollama、智谱 AI、ChatGLM），批量翻译选中文献的摘要和标题。
*   **字段转换系列**：
    *   `0ConvertNoteToTableContentAndExtraField.cs`
    *   `0ConvertTranslateTitle2Custom6.cs`
    *   `1TransT转换为Custom6.cs`
    *   **功能**：将笔记或翻译后的标题内容转换并存入 Citavi 的自定义字段（如 Custom6）或表格中，便于管理和引用。


### **`2分组转换`：文献组织体系重构**
用于在 Citavi 的不同组织方式（分组、关键词、分类）之间进行批量转换。
*   **`ConvertCategoriesToGroups_v1.0.cs`**：将分类转换为分组。
*   **`ConvertGroupsToCategories_v1.0.cs`**：将分组转换为分类。
*   **`ConvertGroupsToKeyword_v1.0.cs`**：将分组转换为关键词。
*   **`ConvertKeywordsToCategories_v1.0.cs`**：将关键词转换为分类。
*   **`ConvertTTitle2CitationKeyAndShortTitle.cs`**：根据标题生成引用键和短标题。
*   **`清理分组名称中的特殊字符.cs`**：批量清理分组名中的非法或多余字符，保持整洁。

### **`3Reference管理`：文献条目标准化**
专注于文献条目本身的信息管理和格式化。
*   **引用键格式化**：
    *   `1_重置引用键为系统默认.cs`
    *   `2_重置引用键为简短格式.cs`
    *   `3_重置引用键为自定义格式.cs`
    *   **功能**：提供三种方式批量重置项目中的引用键格式。
*   **信息自动补全与修正**：
    *   `提取DOI链接到DOI字段.cs`：从其他字段中提取 DOI 并填入标准 DOI 字段。
    *   `根据DOI或PMID更新文献信息.cs`：通过 DOI 或 PMID 在线查询并补全文献信息。
    *   `根据ISSN自动补全期刊信息.cs`：通过 ISSN 自动补全期刊的详细信息。
    *   `根据分割符调换修改title.cs`：根据特定分隔符对标题进行批量修改和调整。

### **`4Knowledge管理`：知识卡片高效整理**
对 Citavi 的核心功能——知识模块进行批量操作和高级管理。
*   **`1_导出知识组为Markdown文档.cs`**：将一个知识组下的所有条目导出为结构化的 Markdown 文档。
*   **`1_按文献重排知识并生成子标题.cs`**：将知识条目按其来源文献重新组织，并自动生成文献标题作为子标题。
*   **`2_按文献标题创建分类并同步.cs`**：根据文献标题自动创建分类，并将对应的知识条目归入其中。
*   **`3_合并多个知识分类到指定的一个分类中.cs`**：将多个分类的知识条目合并到一个目标分类下。
*   **`4_检测隐藏且没有entity的knowledge并删除.cs`**：清理无效的“僵尸”知识条目。
*   **`5_检测隐藏的Annotation并删除.cs`**：清理无效的隐藏标注。
*   **`Knowledge界面预览PDF高亮.cs`**：在知识管理界面直接预览 PDF 上的高亮区域。
*   **`从剪切版获取条目 Open KnowledgeItem with ID from Clipboard.cs`**：通过剪贴板中的 ID 快速打开对应的知识条目。


### **`5附件信息修整`：附件管理与修复**
解决 PDF 等附件的路径问题，确保链接正常。
*   **`1重新定位附件路径.cs`**：当附件移动后，批量更新其路径链接。
*   **`2查找未关联条目的知识项.cs`**：找出那些没有关联到任何文献条目的“孤儿”知识项。
*   **`批量移动附件到Citavi文件夹.cs`**：将散落在各处的附件统一移动到 Citavi 项目的标准附件文件夹中。
*   **`拆分和修正附件路径.cs`**：对不规范的附件路径进行批量修正。
*   **`附件文件夹文件链接状态检查器.cs`**：扫描所有附件，检查链接是否有效，并生成报告。


### **`6PDF导入`：批量导入与自动分组**
高效地将本地 PDF 文件导入 Citavi，并自动进行组织。
*   **`1按文件夹分层并Group.cs`**：导入时，根据源文件夹的层级结构自动创建分组。
*   **`2按文件夹分层并Category.cs`**：导入时，根据源文件夹的层级结构自动创建分类。
*   **`3分文件夹批量导入课件PDF.cs`**：专门用于批量导入存放在不同文件夹中的课件类 PDF。
*   **`4导入单一文件夹内的PDF.cs`**：导入指定单个文件夹内的所有 PDF 文件。


### **`Macros自己的参考` & `z已做成插件`：开发资源与成品**
*   **`Macros自己的参考`**：存放一些宏开发过程中的参考代码、旧版本或被弃用的功能，是学习和二次开发的宝贵资源。
*   **`z已做成插件`**：这部分是宏的“最终形态”。一些功能稳定、使用频繁的宏已经被进一步封装成了独立的 Citavi 插件（`.dll` 文件），提供更稳定、更集成的用户体验。例如字体设置、重复项查找合并、高级搜索等功能。

### **根目录：可能不太好分类的**
这些宏不依赖特定场景，是日常使用中的“快捷方式”和“瑞士军刀”。
*   **`FromExcel导入分组信息.cs`**：从 Excel 表格中读取分组信息，并批量应用到 Citavi 项目中，适合大规模整理。
*   **`OpenBin打开回收站.cs`**：提供一个快捷入口，一键打开 Citavi 的回收站，方便恢复误删项目。
*   **`Preview预览MD-web.cs`**：可能用于在 Citavi 中或通过浏览器预览 Markdown 格式的内容。
*   **`UseBookxNotePro打开PDF附件.cs`**：调用 BookxNotePro 软件打开当前选中文献的 PDF 附件，实现深度阅读和笔记联动。
*   **`UsePDFXchange打开PDF附件.cs`**：调用 PDF-XChange Editor 打开 PDF 附件，利用其强大的标注和编辑功能。
*   **`Use沉浸式阅读翻译PDF.cs`**：结合沉浸式翻译等工具，一键对当前 PDF 进行翻译，提升外文文献阅读效率。
*   **`模板：获取Knowledge和Annotation.cs`**：一个开发模板，用于获取知识条目和标注信息，方便二次开发。
*   **`模板：获取Reference.Id.cs`**：另一个开发模板，用于获取当前选中文献的 ID。
*   **`统一修改摘要和目录字体.cs`**：批量修改项目中所有摘要和目录的字体样式，保持格式统一。

总而言之，这套宏集合从**自动化、跨平台联动、数据清洗、知识整理**等多个维度，全方位地增强了 Citavi 的能力，是一套非常成熟和强大的生产力工具。
