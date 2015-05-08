using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Excel2Git
{
    using Octokit;

    static class GitRepo
    {
        public static int ImportXls(string xlsPath, Uri repoUri)
        {
            int recordsImported = 0;

            // test xls file
            // test repo name

            if (File.Exists(xlsPath))
            {
                var client = new GitHubClient(new ProductHeaderValue("mspnp-importer"), repoUri);

                var user = client.User.Current();



            }
            else
            {
                Console.WriteLine("Xls file not found [{0}]", xlsPath);
            }

            return recordsImported;
        }
    }
}
