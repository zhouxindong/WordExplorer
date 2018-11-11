using System.IO;
using System.Windows.Forms;

namespace WordExplorer
{
    public class Dialogs
    {
        /// <summary>
        /// This class is purely static.
        /// </summary>
        private Dialogs() { }

        /// <summary>
        /// Stores the last accessed file path, so that the next file dialog session should start from there.
        /// </summary>
        private static string gDocumentPath = Application.StartupPath;
        /// <summary>
        /// Stores the name of the currently loaded dopcument. Name is stored without extension to be used when saving in different file format.
        /// </summary>
        private static string gDocumentName;

        /// <summary>
        /// Selects file name for document to open.
        /// </summary>
        /// <returns>File name of the document selected to be opened.</returns>
        public static string OpenDocument()
        {
            return OpenDocument(Globals.OpenFileFilter);
        }

        /// <summary>
        /// Selects AsposePdf XML file name for document to open.
        /// </summary>
        /// <returns>File name of the document selected to be opened.</returns>
        public static string OpenAsposePdfDocument()
        {
            return OpenDocument(Globals.OpenAsposePdfFileFilter);
        }

        /// <summary>
        /// Selects file name for document to open.
        /// </summary>
        /// <param name="filter">File name filter.</param>
        /// <returns>File name of the document selected to be opened.</returns>
        public static string OpenDocument(string filter)
        {
            // Optimized to allow automatic conversion to VB.NET
            OpenFileDialog dlg = new OpenFileDialog();
            try
            {
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                dlg.Title = Globals.OpenDocumentDialogTitle;
                dlg.InitialDirectory = gDocumentPath;
                dlg.Filter = filter;

                // Optimized to allow automatic conversion to VB.NET
                DialogResult dlgResult = dlg.ShowDialog(Globals.MainForm);
                if (dlgResult.Equals(DialogResult.OK))
                {
                    gDocumentPath = Path.GetDirectoryName(dlg.FileName);
                    gDocumentName = Path.GetFileNameWithoutExtension(dlg.FileName);
                    return dlg.FileName;
                }
                else
                {
                    return string.Empty;
                }
            }
            finally
            {
                dlg.Dispose();
            }
        }

        /// <summary>
        /// Selects file name for saving currently opened document.
        /// </summary>
        /// <returns>File name to save the currently chosen document with.</returns>
        public static string SaveDocument()
        {
            return SaveDocument(Globals.SaveFileFilter);
        }

        /// <summary>
        /// Selects PDF file name for saving currently opened document.
        /// </summary>
        /// <returns>File name to save the currently chosen document with.</returns>
        public static string SavePdfDocument()
        {
            return SaveDocument(Globals.SavePdfFileFilter);
        }

        /// <summary>
        /// Selects file name for saving currently opened document.
        /// </summary>
        /// <param name="filter">File name filter.</param>
        /// <returns>File name to save the currently chosen document with.</returns>
        public static string SaveDocument(string filter)
        {
            // Optimized to allow automatic conversion to VB.NET
            SaveFileDialog dlg = new SaveFileDialog();
            try
            {
                dlg.CheckFileExists = false;
                dlg.CheckPathExists = true;
                dlg.Title = Globals.SaveDocumentDialogTitle;
                dlg.InitialDirectory = gDocumentPath;
                dlg.Filter = filter;
                dlg.FileName = gDocumentName;

                // Optimized to allow automatic conversion to VB.NET
                DialogResult dlgResult = dlg.ShowDialog(Globals.MainForm);
                if (dlgResult.Equals(DialogResult.OK))
                {
                    gDocumentPath = Path.GetDirectoryName(dlg.FileName);
                    return dlg.FileName;
                }
                else
                {
                    return string.Empty;
                }
            }
            finally
            {
                dlg.Dispose();
            }
        }
    }
}