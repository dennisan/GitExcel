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
            string repoOwner = "";
            string username = "";
            string password = "";

            if (args.Length < 2)
            {
                Usage();
                return;
            }

            xlsFile = args[0];
            repoName = args[1];
            repoOwner = args[2];

            Console.Write("Enter Git username: ");
            username = Console.ReadLine();
            Console.Write("Enter Git password: ");
            password = Console.ReadLine();
            Console.WriteLine("");

            Console.WriteLine("Importing backlog items from {0} to Git Repo {1}", xlsFile, repoName);

            var repo = new GitRepo(username, password);
            Task<int> t = repo.ImportXls(xlsFile, repoName, repoOwner);
            t.Wait();

            Console.WriteLine("{0} issues imported in the Git repo {1}", t.Result, repoName);

        }

        static void Usage()
        {
            Console.WriteLine("Excel2Git.exe - Utility to import issues from an Excel wooksheet to a Git repository");
            Console.WriteLine("Usage:  Excel2Git.exe xlsfile repo <username> <password>");
            Console.WriteLine("  xlsfile:   Path to the xls file to import.");
            Console.WriteLine("  Repository Name:  Name of the Git repository where issues should be import");
            Console.WriteLine("  Repository Owner: Owner of the Git repository where issues should be import");
        }
    }
}
