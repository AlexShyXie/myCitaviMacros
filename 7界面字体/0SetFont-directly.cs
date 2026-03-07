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
using System.Drawing;
//by qq 1429851250
//插件交流群 202835945

public static class CitaviMacro
{
	public static void Main()
	{
	
		Project project = Program.ActiveProjectShell.Project;

		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;


	    Font font = new Font(mainForm.Font.FontFamily,12); // 在此处指定所需的字体名称和字体大小
        mainForm.Font = font;
		// MessageBox.Show("设置成功");
		
	}
}