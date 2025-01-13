using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDataverseTool
{
    public static  class EnvTool
    {
        public static string GetExeFolder()
        {
            return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        }

        //the profile name is default.json saved in the exe folder
        public static string GetDefaultProfilePath()
        {
            return System.IO.Path.Combine(GetExeFolder(), "taskconfig.json");
        }
    }
}
