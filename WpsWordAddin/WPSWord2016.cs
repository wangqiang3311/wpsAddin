using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using AddInDesignerObjects;
using Office;
using System.Windows.Forms;
using System.Reflection;
using Word;

using System.Drawing;

namespace WpsWordAddin
{
    public class WPSWord2016 : IDTExtensibility2, IRibbonExtensibility
    {
        // Fields
        public static Word.Application wpsapp=null;

        public static object app;

        private void wpsDo()
        {
            object missing = Type.Missing;
            object newTemplate = false;
            object documentType = 0;
            object visible = true;
            var document = wpsapp.Documents.Add(ref missing, ref newTemplate, ref documentType, ref visible);
            wpsapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wpsapp.Selection.Range.Text = "wps hello,world";
            document.Shapes.AddPicture("http://img.kingsoft.com/publish/kingsoft/images/gb/sy/logo.gif", 100, 50, 0x94, 60, ref missing, ref missing, ref missing);
        }
        public string GetCustomUI(string RibbonID)
        {
            return Resource1.myRibbon;
        }

        public Bitmap GetRibbonImage(IRibbonControl ctrl)
        {
            switch (ctrl.Id)
            {
                case "Mu_TYCS":
                    return Resource1.gdc_logo;

                case "Bt_TYABOUT":
                    return Resource1.gingerdroid;
            }
            return null;
        }

        public void GetRibbonOnAction(IRibbonControl ctrl)
        {
            string id = ctrl.Id;
            if (id != null)
            {
                if (id == "Bt_TYCS1")
                {
                    if (wpsapp != null)
                        wpsDo();

                }
                if (id == "Bt_TYABOUT")
                {
                    MessageBox.Show("C#开发word");
                }
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {

        }
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {

        }
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            app = Application;
            Console.WriteLine("插件已连接");
        }
        public void OnStartupComplete(ref Array custom)
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes)
            {
                if (process.ProcessName.ToLower() == "wps")
                {
                    Console.WriteLine(process.ProcessName + "进程已启动");

                    wpsapp = app as Word.Application;
                    wpsapp.DocumentBeforeClose += new ApplicationEvents4_DocumentBeforeCloseEventHandler(wpsapp_DocumentBeforeClose);

                    break;
                }
            }
        }
        void wpsapp_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            Console.WriteLine(Doc.FullName + "wps文档将要关闭");
        }
    }
}
