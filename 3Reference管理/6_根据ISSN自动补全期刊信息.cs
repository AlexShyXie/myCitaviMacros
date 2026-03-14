using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Net;
using System.Linq;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Persistence;
// autoRef Newtonsoft.Json;
using Newtonsoft.Json;
// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.
// 根据文献的ISSN号，从本地或在线数据库自动获取并补全期刊的名称、缩写等信息。
public static class CitaviMacro
{
	public class Journal
	{
	    [JsonProperty("JournalId")]
	    public int JournalId { get; set; }

	    [JsonProperty("JournalTitle")]
	    public string JournalTitle { get; set; }

	    [JsonProperty("MedAbbr")]
	    public string MedAbbr { get; set; }

	    [JsonProperty("ISSN (Print)")]
	    public string IssnPrint { get; set; }

	    [JsonProperty("ISSN (Online)")]
	    public string IssnOnline { get; set; }

	    [JsonProperty("IsoAbbr")]
	    public string IsoAbbr { get; set; }

	    [JsonProperty("NlmId")]
	    public string NlmId { get; set; }
	}
		
	public static void Main()
	{
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		bool readFromWeb = false;
		string json;
		
		if (readFromWeb)// 从网页读取
		{
	        //string completeList;
	        string url = "https://ftp.ncbi.nih.gov/pubmed/J_Entrez.txt";
	        List<string> completeList = new List<string>();
			
			// 下载数据
	        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
	        request.Method = "GET";

	        // 忽略不受信任的证书
	        request.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

	        try
	        {
	            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
	            {
	                if (response.StatusCode == HttpStatusCode.OK)
	                {
	                    Encoding encoding = Encoding.UTF8;
	                    using (Stream stream = response.GetResponseStream())
	                    using (StreamReader reader = new StreamReader(stream, encoding))
	                    {
	                        while (!reader.EndOfStream)
	                        {
	                            string line = reader.ReadLine();
	                            completeList.Add(line);
	                        }
	                    }

	                    DebugMacro.WriteLine("下载完成");
	                }
	                else
	                {
	                    DebugMacro.WriteLine("HTTP响应状态码："+response.StatusCode);
	                }
	            }
	        }
	        catch (Exception ex)
	        {
	            Console.WriteLine(ex.Message);
	        }
			
	        // 将列表转换为单个字符串
	        json = string.Join(Environment.NewLine, completeList);
	        // 在这里处理完整列表
	        // foreach (string line in completeList)
	        // {
	        //     DebugMacro.WriteLine(line);
	        // }
		
			
		}
		else
		{
			// 从文件里读取已经json化的journal信息，读取 JSON 文件内容
			json = File.ReadAllText(@"G:\OneDrive - xiehui1573\Appdata_my\各种软件Setting样式\Citavi重复\journals.json");
			
        	//Journal[] journals = JsonConvert.DeserializeObject<Journal[]>(json);
		}
		// 解析 JSON 数据并转换为对象
		Journal[] myjournals = JsonConvert.DeserializeObject<Journal[]>(json);
        // 处理读取到的数据
        //foreach (Journal journal in myjournals.Skip(20507).Take(5).ToList())
        // {
		// 	DebugMacro.WriteLine("Journal ID: " + journal.JournalTitle);
		// }
		
		// foreach (Journal j in myjournals)
		// {
		// 	if (j.JournalTitle == "Frontiers in oncology")
		// 	{	
		// 		DebugMacro.WriteLine("JournalTitle: " + j.IssnOnline);
		// 	}
		// }

		//if this macro should ALWAYS affect all titles in active project, choose:
		//ProjectReferenceCollection references = project.References;		

		//if this macro should affect just filtered rows in the active MainForm, choose:
		List<Reference> references = mainForm.GetSelectedReferences();	
		List<Periodical> journalCollection = project.Periodicals.ToList();
		
		// 获取项目中的所有Periodical
	
		foreach (Reference reference in references)
		{
			// your code
			
			DebugMacro.WriteLine(reference.Periodical.Name);
			// 查询ISSNOnline等于"2234-943X"的Journal对象，并返回其JournalTitle属性
	        Journal thisjournal = (from j in myjournals
	                                where j.IssnOnline == reference.Periodical.Name
	                                select j).FirstOrDefault();
			if (thisjournal == null)
			{
				thisjournal = (from j in myjournals
	                                where j.IssnPrint == reference.Periodical.Name
	                                select j).FirstOrDefault();
			}
				
	        // 调用 IsinPeriodicals 静态方法，并将 PeriodicalHelper 类名作为前缀
			PeriodicalHelper yesPeriodical = new PeriodicalHelper(); 
	        Periodical newPeriodical = yesPeriodical.IsinPeriodicals(project,journalCollection,thisjournal);
			reference.Periodical = newPeriodical;
			
		}
		

	}
	
	public class PeriodicalHelper
	{
	    public Periodical IsinPeriodicals(Project project, List<Periodical> journalCollection, Journal myjournal)
	    {
			string JournalName = myjournal.JournalTitle;
	        Periodical thisjournal =new Periodical(project, myjournal.JournalTitle); // 将初始值设置为 null

	        foreach (Periodical periodical in journalCollection)
	        {
	            if (JournalName.ToLower() == periodical.Name.ToLower())
	            {
					DebugMacro.WriteLine("找到匹配项。JournalTitle: " + periodical.Name);
					
					periodical.StandardAbbreviation = myjournal.MedAbbr;
					periodical.UserAbbreviation1 = myjournal.IsoAbbr;
					periodical.Issn = myjournal.IssnOnline;
					periodical.Eissn = myjournal.IssnPrint;
	                thisjournal = periodical;
	                break;  // 找到匹配项后跳出循环
	            }
	        }
			if (string.IsNullOrEmpty(thisjournal.Issn))
			{
				DebugMacro.WriteLine("没有找到匹配项。手动添加中...");
				thisjournal.Name = myjournal.JournalTitle;
				thisjournal.StandardAbbreviation = myjournal.MedAbbr;
				thisjournal.UserAbbreviation1 = myjournal.IsoAbbr;
				thisjournal.Issn = myjournal.IssnOnline;
				thisjournal.Eissn = myjournal.IssnPrint;
			}
	        return thisjournal;  // 返回找到的期刊，如果找不到则返回 null
	    }
	}
}