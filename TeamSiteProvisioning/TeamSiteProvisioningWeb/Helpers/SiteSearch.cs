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

            return this.GetSiteCollectionsInPath("teams");
        }

        private List<string> GetSiteCollectionsInPath(string path)
        {
            var siteUri = "https://rdoyle-admin.sharepoint.com/";
            var baseUri = "https://rdoyle.sharepoint.com/";

            var results = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri))
            {

                var tenant = new Tenant(context);
                var spp = tenant.GetSiteProperties(0, true);

                context.Load(spp);
                context.ExecuteQuery();

                foreach (var site in spp)
                {
                    var uri = site.Url.ToString();
                    if (!uri.StartsWith(baseUri))
                    {
                        continue;
                    }

                    uri = uri.Remove(0, baseUri.Length);
                    var uriParts = uri.Split('/').ToArray();
                    if (uriParts.Any() && uriParts[0] == path)
                    {
                        results.Add(site.Url.ToString());
                    }
                }
            }

            return results;
        }

        private List<string> GetSiteEvents()
        {
            var siteUri = "https://rdoyle.sharepoint.com/sites/test01";
            var events = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var lists = context.LoadQuery(context.Web.Lists.Where(l => l.Title == "events"));
                context.ExecuteQuery();
                foreach (var list in lists)
                {
                    var qry = CamlQuery.CreateAllItemsQuery();
                    var items = list.GetItems(qry);
                    context.Load(items);
                    context.Load(items, icol => icol.Include(i => i.DisplayName));
                    context.ExecuteQuery();
                    events.AddRange(items.Select(i => i.DisplayName));
                }

                return events;
            }
        }

        private List<string> GetSiteTasks()
        {
            var siteUri = "https://rdoyle.sharepoint.com/sites/test01";
            var tasks = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var lists = context.LoadQuery(context.Web.Lists.Where(l => l.Title == "tasks"));
                context.ExecuteQuery();
                foreach (var list in lists)
                {
                    var qry = CamlQuery.CreateAllItemsQuery();
                    var items = list.GetItems(qry);
                    context.Load(items, icol => icol.Include(i => i.DisplayName));
                    context.ExecuteQuery();
                    tasks.AddRange(items.Select(i => i.DisplayName));
                }

                return tasks;
            }
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

        private List<string> GetSiteDocuments()
        {
            var siteUri = "https://rdoyle.sharepoint.com/";
            var docs = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri))
            {
                var results = new Dictionary<string, IEnumerable<File>>();
                var lists = context.LoadQuery(context.Web.Lists.Where(l => l.BaseType == BaseType.DocumentLibrary));
                context.ExecuteQuery();
                foreach (var list in lists)
                {
                    var qry = new CamlQuery();
                    qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
                    var items = list.GetItems(qry);
                    context.Load(items, icol => icol.Include(i => i.File));
                    results[list.Title] = items.Select(i => i.File);
                }
                context.ExecuteQuery();

                foreach (var result in results)
                {
                    // Filter by just the documents list
                    if (result.Key.ToString() == "Documents")
                    {
                        foreach (var doc in result.Value)
                        {
                            docs.Add(doc.Name);
                        }
                    }
                }
            }

            return docs;
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