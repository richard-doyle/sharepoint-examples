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
        private ContextFactory contextFactory;

        public SiteProvisioner()
        {
            this.contextFactory = new ContextFactory();
        }

        public void ProvisionSite(SiteDetails details)
        {
            Uri siteUri = new Uri(ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"]);

            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            string accessToken = TokenHelper.GetAppOnlyAccessToken(
                TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

            // Create Site
            var newSiteUri = this.CreateSite(details);
            // Activate Site Publishing
            // These only seems to work with FeatureDefinitionScope.None, rather than FeatureDefinitionScope.Site or FeatureDefinitionSite.Web
            Guid publishingSiteGuid = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
            Guid publishingWebGuid = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
            this.ActivateFeatureOnSite(newSiteUri.ToString(), publishingSiteGuid, FeatureDefinitionScope.None);
            this.ActivateFeatureOnWeb(newSiteUri.ToString(), publishingWebGuid, FeatureDefinitionScope.None);
        }

        private string CreateSite(SiteDetails siteDetails)
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
            using (var adminContext = this.contextFactory.GetContext(tenantAdminUri.ToString()))
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

            return webUrl.ToString();
        }

        private void ActivateFeatureOnSite(string webUrl, Guid featureId, FeatureDefinitionScope scope)
        {
            using (var context = this.contextFactory.GetContext(webUrl))
            {
                var features = context.Site.Features;
                context.Load(features);
                context.ExecuteQuery();

                features.Add(featureId, true, scope);
                context.ExecuteQuery();
            }
        }

        private void ActivateFeatureOnWeb(string webUrl, Guid featureId, FeatureDefinitionScope scope)
        {
            using (var context = this.contextFactory.GetContext(webUrl))
            {
                var features = context.Web.Features;
                context.Load(features);
                context.ExecuteQuery();

                features.Add(featureId, true, scope);
                context.ExecuteQuery();
            }
        }
    }
}