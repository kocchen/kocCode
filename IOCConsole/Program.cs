using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BLL;
using IServer;
using System.Reflection;
using System.Configuration;
using Microsoft.Office.Tools.Word;
using Microsoft.Practices.Unity;
using Microsoft.Office.Interop.Word;

namespace IOCConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //IPressWator pw = (IPressWator)Assembly.Load(ConfigurationManager.AppSettings["AssemName"]).CreateInstance(ConfigurationManager.AppSettings["WaterToolName"]);
            //IPressWator p = new PressWatorBLL();
            //p.GetWatorTools().GetWator();
            //pw.GetWatorTools().GetWator();


             //UnityContainer container = new UnityContainer();
            //container.RegisterType<IServer.IPressWator,BLL.PressWatorBLL>();
            //IServer.IPeople people = container.Resolve<BLL.PeopleBLL>();
            //people.toDoSomething();
            //Console.ReadLine();
            //string strTempContent = "1.168 Agency Background";
            //strTempContent = System.Text.RegularExpressions.Regex.Replace(strTempContent, @"[^\d\.]*", "");
            //Console.WriteLine(strTempContent);
            //Console.ReadLine();
            //dynamic test=new DynamicTest();
            //Console.WriteLine(test.Add(1, 2));

            //var t = new DynamicTest();
            //var method = t.GetType().GetMethod("Add");
            //Console.WriteLine(method.Invoke(t, new object[] {1, 2}));
            
            //Console.ReadLine();

            //var a = new Program();
            //var c = a.T2();
            //Console.WriteLine(c.A);
            //Console.ReadLine();
              _Application wordApp = null;
            wordApp = new Application();
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Visible = false;
         _Document wordDoc = null;
              Object path =  "E:\\workPlace\\source\\Solution Platform\\Proposal Tool\\Proposal Tool\\RFPDocuments\\Asset Management\\2. Terms and Conditions\\test.docx";
              Object template = "E:\\workPlace\\source\\Solution Platform\\Proposal Tool\\Proposal Tool\\RFPDocuments\\Asset Management\\1. Introduction\\1.6 Pre-Proposal Meeting.docx";
            object missing = System.Reflection.Missing.Value;
            wordDoc = wordApp.Documents.Open(ref template, ref missing,
             ref missing, ref missing, ref missing, ref missing, ref missing,
             ref missing, ref missing, ref missing, ref missing, ref missing,
             ref missing, ref missing, ref missing, ref missing);
            foreach (Microsoft.Office.Interop.Word.ContentControl cc in wordDoc.ContentControls)
            {
                
                if (cc.Title == "Pre-Proposal Meeting Time")
                {
                     cc.DropdownListEntries[5].Select();
                    var t = cc.DropdownListEntries[1];
                    //foreach (ContentControlListEntry entry in cc.DropdownListEntries)
                    //{

                    //    //entry.Value = "2a";
                    //    //cc.Range.Text = "2a";
                    //}
                    //DropDownListContentControl contentControl = (cc as DropDownListContentControl);
                    //int m = ((DropDownListContentControl) cc).DropDownListEntries.Count;
                    //ContentControlListEntries list = (cc as DropDownListContentControl).DropDownListEntries;
                    //cc.Range.Text = "1a";
                }
            }

            //Range myRange = wordDoc.Range(0, 0);
            //myRange.InsertBefore("tessst");
            //wordApp.Selection.TypeParagraph();
            //object oStyleName = "Heading 1";
            //Range myRange1 = wordDoc.Range(0, 0);
            //myRange1.set_Style(ref oStyleName);

            //wordDoc.SaveAs(ref path, ref missing, ref missing,
            //   ref missing, ref missing, ref missing, ref missing,
            //   ref missing, ref missing, ref missing, ref missing,
            //   ref missing, ref missing, ref missing, ref missing,
            //   ref missing);
            wordDoc.Close();
            if (wordApp != null)
            {
                //Quit
                wordApp.Quit();
                wordApp = null;
            }
           
        }
        public  dynamic T2()
        {
            var n = 999;
            dynamic result = new { A = n };
            n = 10;
            return result;
        }
    }

    class DynamicTest
    {
        public int Add(int a, int b)
        {
            return a + b;
        }
    }
}
