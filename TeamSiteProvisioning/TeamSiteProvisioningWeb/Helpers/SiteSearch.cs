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
            var siteUri = new Uri(ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"]);
            var resultsList = new List<string>();

            using (var context = this.contextFactory.GetContext(siteUri.ToString()))
            {
                var keywordQuery = new KeywordQuery(context);
                keywordQuery.QueryText = "ContentType=\"Site Collection Metadata\"";
                keywordQuery.TrimDuplicates = false;
                var searchExecutor = new SearchExecutor(context);
                var results = searchExecutor.ExecuteQuery(keywordQuery);
                context.ExecuteQuery();

                var result = results.Value[0];
                foreach (var res in result.ResultRows) {
                    var id = res["UniqueId"];
                    var siteTitle = res["SiteName"];
                    if (this.IsProjectSite(Guid.Parse(id.ToString()), siteTitle.ToString(), "Site Metadata"))
                    {
                        resultsList.Add(siteTitle.ToString());
                    }
                }
            }

            return resultsList;
        }

        private bool IsProjectSite(Guid itemId, string siteUri, string listTitle)
        {
            using (var context = this.contextFactory.GetContext(siteUri))
            {
                var site = context.Web;
                var list = site.Lists.GetByTitle(listTitle);

                var query = new CamlQuery();
                query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='UniqueId'/><Value Type='Guid'>{0}</Value></Contains></Where></Query></View>", itemId);
                var collListItem = list.GetItems(query);

                context.Load(collListItem);
                context.ExecuteQuery();

                return collListItem[0].FieldValues["Site_x0020_Category"].ToString() == "Project";
            }
        }
    }
}