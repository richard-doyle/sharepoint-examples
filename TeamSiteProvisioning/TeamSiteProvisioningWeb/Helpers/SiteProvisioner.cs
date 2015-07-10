using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
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
            //Guid publishingSiteGuid = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");
            //Guid publishingWebGuid = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
            //this.ActivateFeatureOnSite(newSiteUri.ToString(), publishingSiteGuid, FeatureDefinitionScope.None);
            //this.ActivateFeatureOnWeb(newSiteUri.ToString(), publishingWebGuid, FeatureDefinitionScope.None);

            // Add Publishing Home Page
            //var newPageName = this.AddPublishingPage(newSiteUri.ToString(), "Home.aspx");
            // Make the page the home page
            //this.SetHomePage(newSiteUri, newPageName);
            // Create Metadata List
            //var newListTitle = this.CreateMetaDataList(newSiteUri);
            // Add MetaData item
            //this.AddMetaData(newSiteUri, newListTitle);
        }

        private string CreateSite(SiteDetails siteDetails)
        {
            string tenantStr = ConfigurationManager.AppSettings["SiteCollectionRequests_SiteUrl"];
            tenantStr = tenantStr.ToLower().Replace("-my", "").Substring(8);
            tenantStr = tenantStr.Substring(0, tenantStr.IndexOf("."));

            var webUrl = String.Format("https://{0}.sharepoint.com/{1}/{2}", tenantStr, "community", siteDetails.Title);
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

        private string AddPublishingPage(string webUrl, string pageName)
        {
            using (var context = this.contextFactory.GetContext(webUrl))
            {
                var webSite = context.Web;
                context.Load(webSite);

                var publishingWeb = PublishingWeb.GetPublishingWeb(context, webSite);
                context.Load(publishingWeb);
                
                if (publishingWeb != null)
                {
                    var pages = context.Site.RootWeb.Lists.GetByTitle("Pages");
                    var existingPages = pages.GetItems(CamlQuery.CreateAllItemsQuery());
                    context.Load(existingPages, items => items.Include(item => item.DisplayName).Where(obj => obj.DisplayName == pageName));
                    context.ExecuteQuery();

                    // Check that page does not already exists
                    if (existingPages == null || existingPages.Count == 0)
                    {
                        // Get Publishing Page Layouts
                        var publishingLayouts = context.Site.RootWeb.Lists.GetByTitle("Master Page Gallery");
                        var allItems = publishingLayouts.GetItems(CamlQuery.CreateAllItemsQuery());
                        context.Load(allItems, items => items.Include(item => item.DisplayName).Where(obj => obj.DisplayName == "BlankWebPartPage"));
                        context.ExecuteQuery();

                        var layout = allItems.Where(x => x.DisplayName == "BlankWebPartPage").FirstOrDefault();
                        context.Load(layout);

                        // Create a publishing page
                        PublishingPageInformation publishingPageInfo = new PublishingPageInformation();
                        publishingPageInfo.Name = pageName;
                        publishingPageInfo.PageLayoutListItem = layout;

                        PublishingPage publishingPage = publishingWeb.AddPublishingPage(publishingPageInfo);
                        publishingPage.ListItem.File.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                        publishingPage.ListItem.File.Publish(string.Empty);
                        context.Load(publishingPage);
                        context.Load(publishingPage.ListItem.File, obj => obj.ServerRelativeUrl);
                        context.ExecuteQuery();
                    }
                }
            }

            return pageName;
        }

        private void SetHomePage(string webUrl, string pageName)
        {
            using (var context = this.contextFactory.GetContext(webUrl))
            {
                var webSite = context.Site;
                context.Load(webSite);
                context.ExecuteQuery();

                webSite.RootWeb.RootFolder.WelcomePage = "Pages/" + pageName;
                webSite.RootWeb.RootFolder.Update();
                context.ExecuteQuery();
            }
        }

        private string CreateMetaDataList(string webUrl)
        {
            using (var context = this.contextFactory.GetContext(webUrl))
            {
                // Create List
                var listTitle = "Site Metadata";
                var webSite = context.Web;

                var listCreationInformation = new ListCreationInformation();
                listCreationInformation.Title = listTitle;
                listCreationInformation.TemplateType = (int)ListTemplateType.Announcements; // This does not need to be an announcements template

                var list = webSite.Lists.Add(listCreationInformation);

                context.ExecuteQuery();

                // Set content type - this content type needs to already exist - this one has been created as a custom content type on SharePoint Online's Content Type Hub (/sites/contentTypeHub)
                var contentType = context.Web.ContentTypes.GetById("0x0100B190AA81CFE7C34992F8872353FA9888");
                list.ContentTypes.AddExistingContentType(contentType);
                context.ExecuteQuery();

                // Set new content type as default
                var currentCTOrder = list.ContentTypes;
                context.Load(currentCTOrder);
                context.ExecuteQuery();
                var order = new List<ContentTypeId>();
                foreach (var ct in currentCTOrder)
                {
                    if (ct.Name.ToLower().Equals("Site Collection MetaData".ToLower()))
                    {
                        order.Add(ct.Id);
                    }
                }
                list.RootFolder.UniqueContentTypeOrder = order;
                list.RootFolder.Update();
                list.Update();
                context.ExecuteQuery();

                return listTitle;
            }
        }

        private void AddMetaData(string webUrl, string listTitle)
        {
            using (var context = this.contextFactory.GetContext(webUrl))
            {
                var list = context.Web.Lists.GetByTitle(listTitle);

                var listItemInfo = new ListItemCreationInformation();
                var listItem = list.AddItem(listItemInfo);
                listItem["Title"] = "Metadata";
                listItem["Site Category"] = "Project";
                listItem.Update();

                context.ExecuteQuery();
            }
        }
    }
}