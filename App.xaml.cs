using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAppWPF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            try
            {
                Current.Resources["Excel"] = new Excel.Application();
            }
            catch
            {
                Environment.Exit(0);
            }
        }
        private void Application_Exit(object sender, EventArgs e)
        {
            (Current.Resources["Excel"] as Excel.Application).Quit();
        }
    }
}
