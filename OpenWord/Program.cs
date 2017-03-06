using BIMTClassLibrary.Controller;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using MSWord = Microsoft.Office.Interop.Word;
namespace OpenWord
{
    class Program
    {
        
        private Microsoft.Office.Interop.Word._Document oDoc;
        private Microsoft.Office.Interop.Word._Application oWordd = new Microsoft.Office.Interop.Word.Application();
        private object oMissing = System.Reflection.Missing.Value;
        static void Main(string[] args)
        {
            try
            {
                System.Diagnostics.Process.Start("winword.exe");//Word
                SmartUwriteLoadController controller = new SmartUwriteLoadController();
                controller.InstallSmartUwrite();
            }
            catch (Exception)
            {
            }

        }

        
    }
}
