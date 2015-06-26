using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using TeamSiteProvisioningWeb.Models;

namespace TeamSiteProvisioningWeb.Helpers
{
    public class SiteProvisioner
    {
        public void ProvisionSite(SiteDetails details)
        {
            Uri siteUri = new Uri(ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"]);

            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            string accessToken = TokenHelper.GetAppOnlyAccessToken(
                TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

            using (var context = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken))
            {
                this.CreateSite(context, details);
            }
        }

        private void CreateSite(ClientContext context, SiteDetails siteDetails)
        {
            string tenantStr = ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"];
            tenantStr = tenantStr.ToLower().Replace("-my", "").Substring(8);
            tenantStr = tenantStr.Substring(0, tenantStr.IndexOf("."));

            var webUrl = String.Format("https://{0}.sharepoint.com/{1}/{2}", tenantStr, "sites", siteDetails.Title);
            var tenantAdminUri = new Uri(String.Format("https://{0}-admin.sharepoint.com", tenantStr));
            string realm = TokenHelper.GetRealmFromTargetUrl(tenantAdminUri);
            var token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, tenantAdminUri.Authority, realm).AccessToken;

            var authenticationManager = new AuthenticationManager();

            //using (var adminContext = TokenHelper.GetClientContextWithAccessToken(tenantAdminUri.ToString(), token))
            var username = ConfigurationManager.AppSettings["SiteCollectionRequests_UserName"];
            var password = ConfigurationManager.AppSettings["SiteCollectionRequests_Password"];
            using (var adminContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(tenantAdminUri.ToString(), username, password))
            {
                var tenant = new Tenant(adminContext);
                var properties = new SiteCreationProperties()
                {
                    Url = webUrl,
                    Owner = "rdoyle@rdoyle.onmicrosoft.com",
                    Title = siteDetails.Title,
                    Template = "STS#0",
                    StorageMaximumLevel = 100,
                    UserCodeMaximumLevel = 100
                };

                SpoOperation op = tenant.CreateSite(properties);
                adminContext.Load(tenant);
                adminContext.Load(op, i => i.IsComplete);
                adminContext.ExecuteQuery();

                while (!op.IsComplete)
                {
                    System.Threading.Thread.Sleep(10000);
                    op.RefreshLoad();
                    adminContext.ExecuteQuery();
                }
            }
        }
    }
}