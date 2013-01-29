using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace ArchiveAddin
{
    public partial class Archive
    {
        private const string ArchiveFolderName = "Archive";

        private Outlook.MAPIFolder _archiveFolder;
        private Outlook.Items _sentItems;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ArchiveRibbon();
        }

        private void Archive_Startup(object sender, EventArgs e)
        {
            _sentItems = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items;
            _sentItems.ItemAdd += SentItems_ItemAdd;

            _archiveFolder = Application.Session.DefaultStore.GetRootFolder().Folders[ArchiveFolderName];
        }

        private void SentItems_ItemAdd(object item)
        {
            var mailItem = item as Outlook.MailItem;
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

        private void Archive_Shutdown(object sender, EventArgs e)
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
            this.Startup += new System.EventHandler(Archive_Startup);
            this.Shutdown += new System.EventHandler(Archive_Shutdown);
        }
        
        #endregion
    }
}
