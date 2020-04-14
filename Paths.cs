using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace Doc_Merger
{
    public class Paths
    {
        #region Properties
        /// <summary>
        /// List containing all paths to files
        /// </summary>
        public List<string> AllFilesPath { get; set; }

        /// <summary>
        /// Document to be generated as result
        /// </summary>
        public string FinalDocument { get; set; }

        /// <summary>
        /// If some error occur in some file, the file path should be logged to this log file
        /// </summary>
        public string ErrorLog { get; set; }

        /// <summary>
        /// List containing all files
        /// </summary>
        public List<string> AllFiles;
        #endregion

        #region Constructors
        public Paths(List<string> args)
        {
            AllFilesPath = args;
            AllFiles = GetAllFiles();
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Validates if all paths and files are Ok to start
        /// </summary>
        /// <returns>false if any exception occurs while creating files and directories</returns>
        public bool CreateAndValidatePaths()
        {
            string currentDir = Directory.GetCurrentDirectory();
            string finalPath = currentDir + @"\Results";
            string finalFile = finalPath + @"\Completed\Result.docx";
            string errorFile = finalPath + @"\Errors\errors.txt";

            try
            {
                CreateDirectoryIfNotExist($"{ currentDir }\\Results");
                CreateDirectoryIfNotExist($"{ finalPath }\\Completed");
                CreateDirectoryIfNotExist($"{ finalPath }\\Errors");

                CreateFileIfNotExist(finalFile);
                CreateFileIfNotExist(errorFile);

                ErrorLog = errorFile;
                FinalDocument = finalFile;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// execute the merge
        /// </summary>
        public void Execute()
        {
            var app = new Application();
            var count = 1;

            var finalFile = app.Documents.Open(FinalDocument);
            finalFile.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

            AllFiles.ForEach(file =>
            {
                try
                {
                    finalFile.ActiveWindow.Selection.InsertFile(file);
                    finalFile.Save();
                    Console.WriteLine($"{count} - {file}\n");
                }
                catch (COMException)
                {
                    using (StreamWriter writer = File.CreateText(ErrorLog))
                    {
                        writer.WriteLine($"{file}\n");
                    }
                }
                count++;
            });

            finalFile.Close();
            app.Quit();
        }

        #endregion

        #region Private Methods
        /// <summary>
        /// Gets the files in all indicated directories
        /// </summary>
        /// <returns>a list containing all files</returns>
        private List<string> GetAllFiles()
        {
            var files = new List<string>();

            foreach (var filePath in AllFilesPath)
            {
                var filesInFilePath = Directory.GetFiles(filePath, "*.docx");
                AllFiles.AddRange(filesInFilePath);
            }

            return files;
        }

        /// <summary>
        /// Validate and create a directory
        /// </summary>
        /// <param name="dir"></param>
        private void CreateDirectoryIfNotExist(string dir)
        {
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        /// <summary>
        /// validate and create a file
        /// </summary>
        /// <param name="file"></param>
        private void CreateFileIfNotExist(string file)
        {
            if (!File.Exists(file))
            {
                var createFile = File.Create(file);
                createFile.Close();
            }
        }
        #endregion
    }
}
