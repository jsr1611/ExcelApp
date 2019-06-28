using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp4
{
    class ExcelApplication : WindowsFormsApplicationBase
    {
            private static ExcelApplication _instance = null;
            public ExcelApplication()
            {
                this.IsSingleInstance = true;
                this.ShutdownStyle = ShutdownMode.AfterAllFormsClose;
            }
            public static ExcelApplication Instance
            {
                get
                {
                    if (_instance == null)
                    {
                        _instance = new ExcelApplication();
                    }
                    return _instance;
                }
            }
            protected override void OnCreateMainForm()
            {
                ExcelForm.CreateForm();
            }
            protected override void OnStartupNextInstance(StartupNextInstanceEventArgs eventArgs)
            {
                ExcelForm.CreateForm();
            }
        }
    
}
