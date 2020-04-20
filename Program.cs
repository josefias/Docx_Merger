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
            string allFilePath = @"C:\Users\jppirespereira\Desktop\FinalReport.docx";
            string finalPath = @"C:\Users\jppirespereira\Desktop\FinalReport.pdf";
        
            //CreateFileIfNotExist(finalPath);

            var app = new Application();
            Console.WriteLine("opening docx\n");
            var finalFile = app.Documents.Open(allFilePath);
            Console.WriteLine("opened docx\nSaving as PDF\n");
            finalFile.SaveAs(finalPath, WdSaveFormat.wdFormatPDF);
            Console.WriteLine("PDF created");
            finalFile.Close();
            app.Quit();
            Console.WriteLine("Done");
            Console.ReadKey();


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
