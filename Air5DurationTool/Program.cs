using Air5DurationTool.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Air5DurationTool
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                if (Common.IsAppRunning("Air5DurationTool"))
                {
                    MessageBox.Show("程序正在运行，请先退出！", "提示！", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                INIClass ini = new INIClass(Application.StartupPath + "\\config.ini");
                Globals.Air5ConnectionStr = ini.IniReadValue("conn", "Air5"); 
                Globals.Zone = int.Parse(ini.IniReadValue("zone", "zone"));  

                Application.Run(new FrmMain());
            }
            catch (Exception ex)
            { 
                LogService.WriteErr(ex.Message);
            }
        }
    }
}
