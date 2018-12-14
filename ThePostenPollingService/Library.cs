using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThePostenPollingService
{
    class Library
    {
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            string baseDir = "C:\\ProgramData\\ThePostenSignering\\Logs";
            if (!System.IO.Directory.Exists(baseDir))
            {
                System.IO.Directory.CreateDirectory(baseDir);
            }
            try
            {
                sw = new StreamWriter(baseDir + "\\CurrentLogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ex.Source.ToString().Trim() + ";" + ex.Message.ToString().Trim());
                sw.Flush();
                sw.Close();
            }
            catch
            {

            }
        }

        public static void WriteErrorLog(String Message)
        {
            StreamWriter sw = null;
            string baseDir = "C:\\ProgramData\\ThePostenSignering\\Logs";
            if (!System.IO.Directory.Exists(baseDir))
            {
                System.IO.Directory.CreateDirectory(baseDir);
            }
            try
            {
                sw = new StreamWriter(baseDir + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ":" + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {

            }
        }
    }
}
