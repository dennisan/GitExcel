using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;

namespace GitExcel
{
    using Octokit;
    using Excel;
    using System.Data;
	using ExcelExporter;

    public class GitRepo
    {
        private GitHubClient Client;
        private string Username;
        private readonly char[] delims = { ' ' };

        public GitRepo(string username, string password)
        {
            Username = username;

            Client = new GitHubClient(new ProductHeaderValue("mspnp-importer"));
            var basicAuth = new Credentials(username, password);
            Client.Credentials = basicAuth;
        }

		/// <summary>
		/// Import Issues from spreadsheet to Git Repo
        /// Tag Fields: Category, Priority, Size, Timeframe, Status 
        /// </summary>
        /// <param name="xlsPath">Path to source spreadhseet</param>
        /// <param name="repoName">Name of the target Git repo</param>
        /// <param name="repoOwner">Owner of the target Git repo</param>
        /// <returns>Number of issues successfully imported </returns>
        public async Task<int> ImportXls(string xlsPath, string repoName, string repoOwner)
        {
            int recordsImported = 0;

            if (string.IsNullOrEmpty(repoName) || string.IsNullOrEmpty(repoOwner))
                throw new ArgumentException("repo name or owner is missing");

            if (!File.Exists(xlsPath))
            {
                Console.WriteLine("Xls file not found [{0}]", xlsPath);
                return recordsImported;
            }

            IIssuesClient issuesClient = Client.Issue;

            try
            {

                using (FileStream xlsStream = File.Open(xlsPath, System.IO.FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader xlsReader = ExcelReaderFactory.CreateOpenXmlReader(xlsStream))
                    {
                        xlsReader.IsFirstRowAsColumnNames = true;
                        DataSet workbook = xlsReader.AsDataSet();
                        DataTable worksheet = workbook.Tables["Sheet1"];
                        string lastCategory = String.Empty;

                        foreach (DataRow row in worksheet.Rows)
                        {
                            string category = row["Category"].ToString();
                            string guidance = row["Guidance"].ToString();
                            string description = row["Description"].ToString();
                            string priority = row["Priority"].ToString();
                            string size = row["Size"].ToString();
                            string timeframe = row["Timeframe"].ToString();
                            string status = row["Status"].ToString();
                            string owner = Username;

                            if (category.Length == 0)
                                category = lastCategory;
                            else
                                lastCategory = category;

                            try
                            {
                                if (guidance.Length > 0)
                                {
                                    var newIssue = new NewIssue(guidance);

                                    if (description.Length > 0) newIssue.Body = description;
                                    if (owner.Length > 0) newIssue.Assignee = owner;

                                    if (size.Length > 0) newIssue.Labels.Add(string.Format("Size {0}", size));
                                    if (priority.Length > 0) newIssue.Labels.Add(string.Format("Pri {0}", priority));
                                    if (timeframe.Length > 0) newIssue.Labels.Add(string.Format("Timeframe {0}", timeframe));
                                    if (status.Length > 0) newIssue.Labels.Add(string.Format("Status {0}", status));
                                    if (category.Length > 0) newIssue.Labels.Add(category);

                                    var issue = await issuesClient.Create(repoOwner, repoName, newIssue);
                                    recordsImported++;

                                    // sleep to avoid spam trigger alert
                                    Thread.Sleep(3500);

                                    Console.WriteLine("Inserting \"{0}\"", guidance);

                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error creating new issue - {0}", e.Message);
                            }

                        }  // foreach row

                        xlsReader.Close();
                        Console.WriteLine("");

                    } // using xlsReader

                } // using xlsStream

            }
            catch (Exception e)
            {
                Console.WriteLine("Error importing issues - {0}", e.Message);
            }

            return recordsImported;

        } // ImportXls method

        /// <summary>
        /// Export Issues from Git Repo to spreadsheet
        /// Tag Fields: Category, Priority, Size, Timeframe, Status 
        /// </summary>
        /// <param name="xlsPath">Path to source spreadhseet</param>
        /// <param name="repoName">Name of the target Git repo</param>
        /// <param name="repoOwner">Owner of the target Git repo</param>
        /// <returns>Number of issues successfully imported </returns>
        public async Task<int> ExportXls(string xlsPath, string repoName, string repoOwner)
        {
            int recordsExported = 0;

            if (string.IsNullOrEmpty(repoName)  || string.IsNullOrEmpty(repoOwner))
                throw new ArgumentException("repo name or owner is missing");

			if (File.Exists(xlsPath))
            {
                Console.Write("Spreadsheet already exisits.  Overwrite [Y/N]? ");
                var key = Console.ReadLine();
                if (key == "N" || key == "n")
                    return recordsExported;
            }

            try
            {
                IIssuesClient issuesClient = Client.Issue;
                
                // set the request filter as necessary
                var request = new RepositoryIssueRequest();
                request.State = ItemState.Open;
               
                // get the issues for the repository
                var issues = await issuesClient.GetAllForCurrent();  //.GetAllForRepository(); // repoOwner, repoName); //, request);

                if (issues.Count > 0)
                {
					var exporter = new ExcelExport();

					DataSet workbook = new DataSet();
                    DataTable worksheet = workbook.Tables.Add("Sheet1");
                    DataColumnCollection columns = worksheet.Columns;

                    var NbrCol = columns.Add("Nbr", typeof(Int32));
                    var CatCol = columns.Add("Category");
                    var DesCol = columns.Add("Description");
                    var GuiCol = columns.Add("Guidance");
                    var AssCol = columns.Add("Assignee");
                    var MilCol = columns.Add("Milestone");
                    var StaCol = columns.Add("Status");
                    var SizCol = columns.Add("Size");
                    var PriCol = columns.Add("Pri");
                    var TimCol = columns.Add("Timeframe");
                    var UrlCol = columns.Add("Url");
                            
                    foreach (Issue issue in issues) 
                    {
                        var row = worksheet.NewRow();
                             
                        row.SetField(NbrCol, issue.Number.ToString());
                        row.SetField(UrlCol, issue.Url.AbsoluteUri);

                        if (issue.Title != null)
                            row.SetField(DesCol, issue.Title);
                                
                        if (issue.Body != null)
                            row.SetField(GuiCol, issue.Body);

                        if (issue.Milestone != null)
                            row.SetField(MilCol, issue.Milestone.Title);

                        if (issue.Assignee != null)
                            row.SetField(AssCol, issue.Assignee.Login);

                        foreach (Label label in issue.Labels) 
                        {
                            // split the label into tokens
                            var tokens = label.Name.Split(delims, 2);
							var labelType = tokens[0].ToLower();

                            switch (labelType)
							{
								case "size":
								case "status":
								case "pri":
								case "timeframe":
									if (columns.Contains(labelType))
										row.SetField(columns.IndexOf(labelType), tokens[1]);
									break;

								default:
									row.SetField(CatCol, label.Name);
									break;
							}

                        }  // foreach label
                                
                        worksheet.Rows.Add(row);
						recordsExported++;

                    }  // foreach issue

					exporter.AddSheet(worksheet);
					exporter.ExportTo(xlsPath);

                }  // count > 0
                
            }  // try

            catch (Exception e)
            {
                Console.WriteLine("Error exporting issues - {0}", e.Message);
            }

            return recordsExported;

        } // ExportXls method

    } // GitRepo class
}
