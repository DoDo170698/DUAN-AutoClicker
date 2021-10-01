using System;
using System.Collections.Generic;
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
            var pathFolder = AppDomain.CurrentDomain.BaseDirectory + @"\LoggingFile";
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
