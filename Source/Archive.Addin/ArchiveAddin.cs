using System;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace Archive.Addin
{
	public partial class ArchiveAddin
	{
		private const string PersonalFolderName = "PersonalFolders";
		private const string ArchiveFolderName = "Archive";

		private MAPIFolder _archiveFolder;
		private Items _sentItems;

		protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
		      return new ArchiveRibbon();
		}

		private void ArchiveAddIn_Startup(object sender, EventArgs e)
		{
			_sentItems = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail).Items;
			_sentItems.ItemAdd += SentItems_ItemAdd;

			_archiveFolder = Application.Session.Folders[PersonalFolderName].Folders[ArchiveFolderName];
		}

		private void SentItems_ItemAdd(object item)
		{
			var mailItem = item as MailItem;
			if (mailItem == null) return;
			mailItem.ShowCategoriesDialog();
			if (string.IsNullOrWhiteSpace(mailItem.Categories))
			{
				mailItem.Delete();
			}
			else
			{
				mailItem.Move(_archiveFolder);
			}
		}

		private void ArchiveAddIn_Shutdown(object sender, EventArgs e)
		{
			_sentItems.ItemAdd -= SentItems_ItemAdd;
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			Startup += ArchiveAddIn_Startup;
			Shutdown += ArchiveAddIn_Shutdown;
		}

		#endregion
	}
}