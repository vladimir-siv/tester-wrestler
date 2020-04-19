using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace WordVSTOWrestler
{
	[ComVisible(true)]
	public class MainRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		#region Ribbon Callbacks
		
		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		public void TestButtonClick(Office.IRibbonControl control)
		{
			MessageBox.Show("Test");
		}

		#endregion

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("WordVSTOWrestler.Ribbons.MainRibbon.MainRibbon.xml");
		}

		#endregion

		#region Helpers

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
