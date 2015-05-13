using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;

namespace Excel2Git
{
    using Octokit;
    using Excel;
    using System.Data;

    public class GitRepo
    {
        private GitHubClient Client;
        private string Username;
        private Uri RepoUri;
        
        public GitRepo(string username, string password)
        {
            Username = username;

            Client = new GitHubClient(new ProductHeaderValue("mspnp-importer"));
            var basicAuth = new Credentials(username, password);
            Client.Credentials = basicAuth;
        }


        public async Task<int> ImportXls(string xlsPath, string repoName, string repoOwner = null)
        {
            int recordsImported = 0;

            if (string.IsNullOrEmpty(repoName))
                throw new ArgumentException("repoName missing");

            if (string.IsNullOrEmpty(repoOwner))
                repoOwner = Username;

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
 
    } // GitRepo class
}
