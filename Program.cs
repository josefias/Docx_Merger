using System;
using System.Linq;
using System.Threading;

namespace Doc_Merger
{
    class Program
    {
        static void Main(string[] args)
        {
            KillAllWordProcesses();

            Paths allPaths;

            if (args.Length <= 0)
            {
                Console.WriteLine("I need some files.\n Run this tool again with the path to the files as arguments");
                Console.ReadKey();
                return;
            }

            allPaths = new Paths(args.ToList());
            if (!allPaths.CreateAndValidatePaths())
            {
                Console.WriteLine("something went wrong!\nBreaking execution...");
                Console.ReadKey();
                return;
            }

            var maxThreads = Environment.ProcessorCount;

            var t1 = new Thread(new ThreadStart(() => allPaths.Execute()));
            var t2 = new Thread(new ThreadStart(() => allPaths.Execute()));

            t1.Start();
            t2.Start();
        }

        private static void KillAllWordProcesses()
        {
            String comand = "/C taskkill /f /IM WINWORD.EXE";
            System.Diagnostics.Process.Start("cmd.exe", comand);
            Console.WriteLine("press any key to start!\n");
            Console.ReadKey();
        }
    }
}
