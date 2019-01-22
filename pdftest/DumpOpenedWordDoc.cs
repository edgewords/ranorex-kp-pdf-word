/*
 * Created by Ranorex
 * User: edgewords
 * Date: 09/01/2019
 * Time: 17:34
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace pdftest
{
    /// <summary>
    /// Description of DumpOpenedWordDoc.
    /// </summary>
    [TestModule("5A322B94-16AC-42DE-9621-1FBD81A70E78", ModuleType.UserCode, 1)]
    public class DumpOpenedWordDoc : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public DumpOpenedWordDoc()
        {
            // Do not delete - a parameterless constructor is required!
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            
//			Word.Application word = null;
//			try
//			{
//			  word = (Word.Application)Marshal.GetActiveObject("Word.Application");
//			}
//			catch (COMException)
//			{
//			}
//			if (word == null) word = new Microsoft.Office.Interop.Word.Application();
//			if(word == null) { /* report error */ }
//			
//			Ranorex.Report.Info("The current document is :" + word.ActiveDocument.FullName.ToString());
            
            
//            Microsoft.Office.Interop.Word.Application WordObj;
//			WordObj = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
//			for (int i = 0; i < WordObj.Windows.Count; i++)
//			{
//			    object idx = i + 1;
//			    var WinObj = WordObj.Windows.get_Item(ref idx);
//			    Ranorex.Report.Info(WinObj.Document.FullName);
//			    
//			}
			
            
			//Get handle to already opened Word doc using marshal service. Both Word & Ranorex must be
			//opened with Administrator permissions
//            
//			Word.Application objWord = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") as Word.Application;
//			
//			string docname = objWord.ActiveDocument.FullName;
//			
//			Ranorex.Report.Info(docname);


			//Get handle on exiting opened Word Doc (need to know name and path in advance)
			Word.Application objWord = new Microsoft.Office.Interop.Word.Application();
			object missing = System.Reflection.Missing.Value;
			Microsoft.Office.Interop.Word.Document document = objWord.Documents.Open(@"C:\Users\edgewords\Documents\ChatLog Meet Now 2019_01_08 22_00.rtf", false, true, false, ref missing, ref missing, false);
			string docname = objWord.ActiveDocument.FullName;
			
			Ranorex.Report.Info(docname);
			

        }
        
        
    }
}
