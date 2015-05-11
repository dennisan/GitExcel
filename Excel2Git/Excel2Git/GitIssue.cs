using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Git
{
    public class GitIssue
    {
        public string Title { get; set; }
        public string body  { get; set; }
        public string Assignee { get; set; }
        public int Milestone { get; set; }
        public List<string> Labels { get; set; }

        public GitIssue(string Title)
        {
             

        }

        async void AddToRepo(string repoName)
        {



        }
    }
}
