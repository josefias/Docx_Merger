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
            Paths paths = new Paths();
            List<string> allFilesPath = default;
            if (args.Length > 0)
                allFilesPath.AddRange(args);

            paths.AllFilePath.AddRange(allFilesPath);
            string currentDir = Directory.GetCurrentDirectory();
            string finalPath = currentDir + @"\Results";
            string finalFile = finalPath + @"\Completed\Result.docx";
            string errorFile = finalPath + @"\Errors\errors.txt";
            ValidatePaths(currentDir, finalPath, finalFile, errorFile);

            var t1 = new Thread(new ThreadStart(() => startup(allFilePath[0], finalPath[0], errorLogPath[0])));
            var t2 = new Thread(new ThreadStart(() => startup(allFilePath[1], finalPath[1], errorLogPath[1])));

            t1.Start();
            t2.Start();
        }

        private static void ValidatePaths(string currentDir, string finalPath, string finalFile, string errorFile)
        {
            if (!Directory.Exists($"{ currentDir }\\Results"))
                Directory.CreateDirectory($"{currentDir}\\Results");

            if (!Directory.Exists($"{ finalPath }\\Completed"))
                Directory.CreateDirectory($"{finalPath}\\Completed");
            if (!Directory.Exists($"{ finalPath }\\Errors"))
                Directory.CreateDirectory($"{finalPath}\\Errors");

            if (!File.Exists(finalFile))
                File.Create(finalFile);
            if (!File.Exists(errorFile))
                File.Create(errorFile);
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
