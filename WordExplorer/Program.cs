﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Aspose.Words;

namespace WordExplorer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var license = new License();
            license.SetLicense(@"Aspose\Aspose.Words.2008.lic");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
