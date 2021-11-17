using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppAutoClick.Helper
{
    public static class LoggingHelper
    {
        public static void Write(string logMessage)
        {
            string pathFileLog = ConfigurationManager.AppSettings["PathFileLog"];
            var pathFolder = string.IsNullOrEmpty(pathFileLog) ? AppDomain.CurrentDomain.BaseDirectory : pathFileLog + @"\LoggingFile";
            bool exists = Directory.Exists(pathFolder);
            if (!exists)
                Directory.CreateDirectory(pathFolder);
            var path = pathFolder + @"\" + $"log_{DateTime.Now.ToString("dd-MM-yyyy")}";
            using (StreamWriter w = File.AppendText(path))
            {
                w.Write("\nLog Entry : ");
                w.Write($"{DateTime.Now.ToString("HH:mm:ss")} {DateTime.Now.ToString("dd-MM-yyyy")}");
                w.Write("  :");
                w.WriteLine($"\r{logMessage}");
                w.WriteLine("-------------------------------");
            }
        }
    }
}
