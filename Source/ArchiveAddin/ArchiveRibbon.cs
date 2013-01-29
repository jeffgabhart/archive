using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ArchiveRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ArchiveAddin
{
    [ComVisible(true)]
    public class ArchiveRibbon : Office.IRibbonExtensibility
    {
        private const string ArchiveFolderName = "Archive";
        private readonly MAPIFolder _archiveFolder;
        private readonly Application _outlook;
        private Office.IRibbonUI _ribbon;

        public ArchiveRibbon()
        {
            _outlook = new Application();
            _archiveFolder = _outlook.Application.Session.DefaultStore.GetRootFolder().Folders[ArchiveFolderName];
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ArchiveAddin.ArchiveRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this._ribbon = ribbonUI;
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

        public void ArchiveButton_Click(Office.IRibbonControl control)
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
