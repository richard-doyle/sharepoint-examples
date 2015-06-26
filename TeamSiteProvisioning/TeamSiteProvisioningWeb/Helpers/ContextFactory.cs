using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace TeamSiteProvisioningWeb.Helpers
{
    public class ContextFactory
    {
        public ClientContext GetContext(string url)
        {
            var authenticationManager = new AuthenticationManager();
            var username = ConfigurationManager.AppSettings["SiteCollectionRequests_UserName"];
            var password = ConfigurationManager.AppSettings["SiteCollectionRequests_Password"];
            return authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(url, username, password);
        }
    }
}