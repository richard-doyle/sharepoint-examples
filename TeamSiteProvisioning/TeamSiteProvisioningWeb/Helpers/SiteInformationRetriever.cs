using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using TeamSiteProvisioningWeb.Models;

namespace TeamSiteProvisioningWeb.Helpers
{
    public class SiteInformationRetriever
    {
        private ContextFactory contextFactory;

        public SiteInformationRetriever()
        {
            this.contextFactory = new ContextFactory();
        }

        public SiteInformation GetInfo()
        {
            var siteUri = "https://rdoyle.sharepoint.com/sites/test01";

            using (var ctx = this.contextFactory.GetContext(siteUri))
            {
                var numDocs = 0;
                var results = new Dictionary<string, IEnumerable<File>>();
                var lists = ctx.LoadQuery(ctx.Web.Lists.Where(l => l.BaseType == BaseType.DocumentLibrary));
                ctx.ExecuteQuery();
                foreach (var list in lists)
                {
                    var items = list.GetItems(CreateAllFilesQuery());
                    ctx.Load(items, icol => icol.Include(i => i.File));
                    results[list.Title] = items.Select(i => i.File);
                }
                ctx.ExecuteQuery();

                foreach (var result in results)
                {
                    // Filter by just the documents list
                    if (result.Key.ToString() == "Documents")
                    {
                        numDocs = result.Value.Count();
                    }
                }   

                return new SiteInformation
                {
                    NumberOfDocuments = numDocs
                };
            }
        }

        private static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
            return qry;
        }
    }
}