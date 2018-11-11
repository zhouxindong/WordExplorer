using Aspose.Words;

namespace WordExplorer
{
    public class Globals
    {
        /// <summary>
        /// This class is purely static, that's why we prevent instance creation by declaring instance constructor as private.
        /// </summary>
        private Globals() { }

        // Titles
        internal static string ApplicationTitle = "Document Explorer";
        internal static readonly string UnexpectedExceptionDialogTitle = ApplicationTitle + " - unexpected error occured";
        internal static readonly string OpenDocumentDialogTitle = "Open Document";
        internal static readonly string SaveDocumentDialogTitle = "Save Document As";

        // File filters
        internal static readonly string OpenFileFilter =
            "All Word Documents (*.docx;*.dotx;*.doc;*.dot;*.htm;*.html;*.rtf;*.xml;*.wml)|*.docx;*.dotx;*.doc;*.dot;*.htm;*.html;*.rtf;*.xml;*.wml|" +
            "Word Documents (*.docx;*.dotx;*.doc;*.dot)|*.docx;*.dotx;*.doc;*.dot|" +
            "All Web Pages (*.htm;*.html)|*.htm;*.html|" +
            "Rich Text Format (*.rtf)|*.rtf|" +
            "XML Files (*.xml;*.wml)|*.xml;*.wml|" +
            "All Files (*.*)|*.*";

        internal static readonly string OpenAsposePdfFileFilter =
            "Aspose.Pdf XML (*.AsposePdf)|*.AsposePdf";

        internal static readonly string SaveFileFilter =
            "Microsoft Word 97 - 2003 Document (*.doc)|*.doc|" +
            "Microsoft Office 2007 Open XML (*.docx)|*.docx|" +
            "Adobe Acrobat Document (*.pdf)|*.pdf|" +
            "Rich Text Format (*.rtf)|*.rtf|" +
            "Microsoft Word 2003 WordprocessingML (*.xml)|*.xml|" +
            "Web Page (*.html)|*.html|" +
            "Plain Text (*.txt)|*.txt|" +
            "Aspose.Pdf XML (*.AsposePdf)|*.AsposePdf";

        internal static readonly string SavePdfFileFilter =
            "Adobe Acrobat Document (*.pdf)|*.pdf";

        /// <summary>
        /// Reference for application's main form.
        /// </summary>
        internal static MainForm MainForm;

        /// <summary>
        /// Reference for currently loaded 鈕cument.
        /// </summary>
        internal static Document Document;
    }

}