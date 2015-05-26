using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitExcel
{
	enum ActionType { Import, Export }; 

    class Program
    {
		ActionType Action = ActionType.Export;
		string XlsFile;
		string Repo;
		string Owner;
		string User;
		string Pass;

		static void Main(string[] args)
        {
			var p = new Program();
	
			if (p.ParseArgs(args))
				p.Run();
		}

		void Run() {

			var repo = new GitRepo(User, Pass);

			if (Action == ActionType.Export)
			{
				Console.WriteLine("Exporting issues from repo {0} to spreadsheet {1}", Repo, XlsFile);
				Task<int> t = repo.ExportXls(XlsFile, Repo, Owner);
				t.Wait();
				Console.WriteLine("{0} issues exported to the spreadsheet {1}", t.Result, XlsFile);
			}
			else
			{
				Console.WriteLine("Importing issues from {0} to Git Repo {1}", XlsFile, Repo);
				Task<int> t = repo.ImportXls(XlsFile, Repo, Owner);
				t.Wait();
				Console.WriteLine("{0} issues imported to the Git repo {1}", t.Result, Repo);

			}

			return;
        }

		bool ParseArgs(string[] args)
		{
			if (args.Length > 0 && args[0].ToLower().Substring(0,2) == "/h")
			{
				Usage();
				return false;
			}

			for (int i = 0; i < args.Length; i++) 
			{
				var arg = args[i].ToLower().Substring(0,2);

				switch (arg)
				{
					case "/i":
						Action = ActionType.Import;
						break;
					case "/e":
						Action = ActionType.Export;
						break;
					case "/r":
						Repo = args[++i];
						break;
					case "/o":
						Owner = args[++i];
						break;
					case "/u":
						User = args[++i];
						break;
					case "/p":
						Pass = args[++i];
						break;
					default:
						XlsFile = args[i];
						break;
				}
			}

			// set owner to user if not set
			if (Owner == null && User != null)
				Owner = User;

			// set outfile to default if exporting
			if (XlsFile == null && Action == ActionType.Export)
				XlsFile = "output.xls";

			if (XlsFile == null || Repo == null || Owner == null || User == null || Pass == null) {
				Usage("Missing one or more required argument(s).");
				return false;
			}

			return true;
		}

        void Usage(string errorMessage = null)
        {
			if (errorMessage != null)
				Console.WriteLine("Error - {0}", errorMessage);

			Console.WriteLine();
			Console.WriteLine("GitExcel - A utility for tranfering issues between a Git repository and an Excel wooksheet.");
			Console.WriteLine();
			Console.WriteLine("Usage: GitExcel <xlsfile> /Help /Import /Export /R <repo> /O <owner> /U <username> /P <password>");
			Console.WriteLine();
			Console.WriteLine("  <xlsfile>    - Full path to the xls file to import or export (Required)");
			Console.WriteLine("  /Import      - Import issues from a spreadsheet to a Git repo");
            Console.WriteLine("  /Export      - Export issues from a Git repo to a spreadsheet (default)");
			Console.WriteLine("  /Help        - Show this help text");
			Console.WriteLine("  <repo>       - Name of the Git repository (Required)");
			Console.WriteLine("  <owner>      - Owner of the Git repository (Required)");
			Console.WriteLine("  <username>   - Git user who has access to the repository (Required)");
			Console.WriteLine("  <password>   - Git user's password (Required)");
        }
    }
}
