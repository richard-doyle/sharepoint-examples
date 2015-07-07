using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace TeamSiteProvisioningWeb.Helpers
{
    public class SiteSearch
    {
        private ContextFactory contextFactory;

        public SiteSearch()
        {
            this.contextFactory = new ContextFactory();
        }

        public List<string> Search(string searchTerm)
        {
            //var username = ConfigurationManager.AppSettings["SiteCollectionRequests_UserName"];
            //this.UserIsMemberOfSite("https://rdoyle.sharepoint.com/sites/test01", "test01 members", username);
            //return this.GetMemberSites();

            return this.GetSiteMembers();
        }

        private List<string> GetMemberSites()
        {
            var siteUri = new Uri(ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"]);

            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var keywordQuery = new KeywordQuery(context);
                keywordQuery.QueryText = "rdoyle AND SiteGroup:\"Team\"";
                keywordQuery.TrimDuplicates = false;
                var searchExecutor = new SearchExecutor(context);
                var results = searchExecutor.ExecuteQuery(keywordQuery);
                context.ExecuteQuery();
            }

            return new List<string>();
        }

        private List<string> GetSiteMembers()
        {
            var siteUri = "https://rdoyle.sharepoint.com/";
            var users = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var groups = context.Web.SiteGroups;
                context.Load(groups);
                context.ExecuteQuery();

                foreach (var group in groups)
                {
                    if (group.Title == "Team Site Members")
                    {
                        var gUsers = group.Users;
                        context.Load(gUsers);
                        context.ExecuteQuery();
                        users = gUsers.Select(u => u.Email).ToList();
                    }
                }
            }

            return users;
        }

        private bool UserIsMemberOfSite(string siteUri, string groupName, string userName)
        {
            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var member = context.Web.IsUserInGroup(groupName, userName);
            }

            return false;
        }

        private List<string> GetHiddenProjectSites()
        {
            var siteUri = new Uri(ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"]);
            var resultsList = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var keywordQuery = new KeywordQuery(context);
                keywordQuery.QueryText = "Project AND HiddenSite AND ContentType=\"Site Collection Metadata\"";
                keywordQuery.TrimDuplicates = false;
                keywordQuery.SelectProperties.Clear();
                keywordQuery.SelectProperties.Add("SiteNumFollowers");
                keywordQuery.SelectProperties.Add("SiteTitle");
                keywordQuery.SelectProperties.Add("SiteMembers");
                var searchExecutor = new SearchExecutor(context);
                var results = searchExecutor.ExecuteQuery(keywordQuery);
                context.ExecuteQuery();

                var result = results.Value[0];
                foreach (var res in result.ResultRows)
                {
                    var siteTitle = res["SiteName"];
                    resultsList.Add(siteTitle.ToString());
                }
            }

            return resultsList;
        }
    }
}