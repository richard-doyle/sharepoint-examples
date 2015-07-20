using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client.WebParts;

namespace TeamSiteProvisioningWeb.Helpers
{
    public class AppPartAdder
    {
        private ContextFactory contextFactory;

        public AppPartAdder()
        {
            this.contextFactory = new ContextFactory();
        }

        public void AddAppPart()
        {
            var siteUri = "https://rdoyle.sharepoint.com/sites/developer";

            using (var context = this.contextFactory.GetContext(siteUri))
            {
                var file = context.Web.GetFileByServerRelativeUrl("Default.aspx");
                var limitedWebPartManager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);

                var xmlWebPart = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<WebPart xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"" +
                " xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"" +
                " xmlns=\"http://schemas.microsoft.com/WebPart/v2\">" +
                "<Title>My Web Part</Title><FrameType>Default</FrameType>" +
                "<Description>Use for formatted text, tables, and images.</Description>" +
                "<IsIncluded>true</IsIncluded><ZoneID></ZoneID><PartOrder>0</PartOrder>" +
                "<FrameState>Normal</FrameState><Height /><Width /><AllowRemove>true</AllowRemove>" +
                "<AllowZoneChange>true</AllowZoneChange><AllowMinimize>true</AllowMinimize>" +
                "<AllowConnect>true</AllowConnect><AllowEdit>true</AllowEdit>" +
                "<AllowHide>true</AllowHide><IsVisible>true</IsVisible><DetailLink /><HelpLink />" +
                "<HelpMode>Modeless</HelpMode><Dir>Default</Dir><PartImageSmall />" +
                "<MissingAssembly>Cannot import this Web Part.</MissingAssembly>" +
                "<PartImageLarge>/_layouts/images/mscontl.gif</PartImageLarge><IsIncludedFilter />" +
                "<Assembly>Microsoft.SharePoint, Version=13.0.0.0, Culture=neutral, " +
                "PublicKeyToken=94de0004b6e3fcc5</Assembly>" +
                "<TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>" +
                "<ContentLink xmlns=\"http://schemas.microsoft.com/WebPart/v2/ContentEditor\" />" +
                "<Content xmlns=\"http://schemas.microsoft.com/WebPart/v2/ContentEditor\">" +
                "<![CDATA[This is a first paragraph!<DIV>&nbsp;</DIV>And this is a second paragraph.]]></Content>" +
                "<PartStorage xmlns=\"http://schemas.microsoft.com/WebPart/v2/ContentEditor\" /></WebPart>";

                var oWebPartDefinition = limitedWebPartManager.ImportWebPart(xmlWebPart);

                limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "Left", 1);

                context.ExecuteQuery();
            }

        }
    }
}