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

            if (args.Length == 2)
            {
                xlsFile = args[0];
                repoName = args[1];
            }

            Console.WriteLine("Importing backlog items from {0} to Git Repo {1}", xlsFile, repoName);

            int recordsImported = GitRepo.ImportXls(xlsFile, new Uri(repoName));

            Console.WriteLine("{0} issues imported in the Git repo {1}", recordsImported, repoName);
        }
    }
}
