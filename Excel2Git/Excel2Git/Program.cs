using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Git
{
    class Program
    {
        static void Main(string[] args)
        {
            string xlsFile = "";
            string repoName = "";
            string username = "";
            string password = "";

            if (args.Length < 2)
            {
                Usage();
                return;
            }
            else
            {
                xlsFile = args[0];
                repoName = args[1];

                if (args.Length == 2)
                {
                    Console.Write("Enter Git username: ");
                    username = Console.ReadLine();
                    Console.Write("Enter Git password: ");
                    password = Console.ReadLine();
                    Console.WriteLine("");
                }
                if (args.Length == 4)
                {
                    username = args[2];
                    password = args[3];
                }
            }



            Console.WriteLine("Importing backlog items from {0} to Git Repo {1}", xlsFile, repoName);

            var repo = new GitRepo(username, password);
            Task<int> t = repo.ImportXls(xlsFile, repoName);
            t.Wait();

            Console.WriteLine("{0} issues imported in the Git repo {1}", t.Result, repoName);

        }

        static void Usage()
        {
            Console.WriteLine("GitImporter.exe - Utility to import issues for and excel wooksheet to a Git repository");
            Console.WriteLine("Usage:  GitImporter.exe xlsfile repo <username> <password>");
            Console.WriteLine("        xlsfile - The path the to xls file to import.");
            Console.WriteLine("        repo - The name of the Git repo where issues should be import.");
            Console.WriteLine("        username - Your Git username.");
            Console.WriteLine("        password - Your Git password.");
        }
    }
}
