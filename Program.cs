using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Doc_Merger
{
    class Program
    {
        static void Main(string[] args)
        {
            killAllWordProcess();
            string[] allFilePath = { @"C:\Users\altavares\Documents\Varian\Docs\sonar result\original\pt3", @"C:\Users\altavares\Documents\Varian\Docs\sonar result\original\pt4" };
            string[] finalPath = { @"C:\Users\altavares\Documents\Varian\Docs\sonar result\FinalReport3.docx", @"C:\Users\altavares\Documents\Varian\Docs\sonar result\FinalReport4.docx" };
            string[] errorLogPath = { @"C:\Users\altavares\Documents\Varian\Docs\sonar result\errorlog3.txt", @"C:\Users\altavares\Documents\Varian\Docs\sonar result\errorlog4.txt" };

            var t1 = new Thread(new ThreadStart(() => startup(allFilePath[0], finalPath[0], errorLogPath[0])));
            var t2 = new Thread(new ThreadStart(() => startup(allFilePath[1], finalPath[1], errorLogPath[1])));

            t1.Start();
            t2.Start();
        }

        private static void startup(string allFilesPath, string finalPath, string errorLogPath)
        {
            var app = new Application();
            var files = Directory.GetFiles(allFilesPath, "*.docx");
            CreateFileIfNotExist(errorLogPath);
            CreateFileIfNotExist(finalPath);

            var finalFile = app.Documents.Open(finalPath);
            var count = 1;
            finalFile.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            files.ToList().ForEach(file =>
            {
                try
                {
                    finalFile.ActiveWindow.Selection.InsertFile(file);
                    finalFile.Save();
                    Console.WriteLine($"{count} - {file}\n");
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    using (StreamWriter writer = File.CreateText(errorLogPath))
                    {
                        writer.WriteLine($"{count} - {file}\n");
                    }
                }
                count++;


            });

            finalFile.Close();
            app.Quit();
        }

        private static void killAllWordProcess()
        {
            String comand = "/C taskkill /f /IM WINWORD.EXE";
            System.Diagnostics.Process.Start("cmd.exe", comand);
            Console.WriteLine("press any key to start!\n");
            Console.ReadKey();
        }

        private static void CreateFileIfNotExist(string finalPath)
        {
            if (!File.Exists(finalPath))
            {
                var file = File.Create(finalPath);
                file.Close();

            }

        }
    }
}
