using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace Archive.Addin
{
	[ComVisible(true)]
	public class ArchiveRibbon : IRibbonExtensibility
	{
		private const string PersonalFolderName = "PersonalFolders";
		private const string ArchiveFolderName = "Archive";

		private readonly MAPIFolder _archiveFolder;
		private readonly Application _outlook;
		private IRibbonUI ribbon;

		public ArchiveRibbon()
		{
			_outlook = new Application();
			_archiveFolder = _outlook.GetNamespace("MAPI").Folders[PersonalFolderName].Folders[ArchiveFolderName];
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("Archive.Addin.ArchiveRibbon.xml");
		}

		#endregion

		#region Ribbon Callbacks

		//Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

		public void ArchiveRibbon_Load(IRibbonUI ribbonUI)
		{
			ribbon = ribbonUI;
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
					using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
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

		public void ArchiveButton_Click(IRibbonControl control)
		{
			if (_outlook.ActiveWindow() == _outlook.ActiveExplorer())
			{
				HandleExplorerWindow();
			}
			else
			{
				HandleEmailWindow();
			}
		}

		private void HandleEmailWindow()
		{
			var item = _outlook.ActiveInspector().CurrentItem as MailItem;
			if (item == null) return;
			item.ShowCategoriesDialog();
			if (string.IsNullOrWhiteSpace(item.Categories)) return;
			if (item.Parent != _archiveFolder)
			{
				item.Move(_archiveFolder);
			}
		}

		private void HandleExplorerWindow()
		{
			Selection items = _outlook.ActiveExplorer().Selection;
			var item = items[1] as MailItem;
			if (item == null) return;
			item.ShowCategoriesDialog();
			if (string.IsNullOrWhiteSpace(item.Categories)) return;
			foreach (MailItem i in items)
			{
				i.Categories = item.Categories;
				if (i.Parent != _archiveFolder)
				{
					i.Move(_archiveFolder);
				}
			}
		}
	}
}