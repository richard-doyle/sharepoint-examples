using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace TeamSiteProvisioningWeb.Helpers
{
    public class ScriptEmbedder
    {
        private ContextFactory contextFactory;

        public ScriptEmbedder()
        {
            this.contextFactory = new ContextFactory();
        }

        public void Embed()
        {
            var siteUri = "https://rdoyle.sharepoint.com/sites/test09";

            using (var context = this.contextFactory.GetContext(siteUri))
            {
                var scriptBlock = "";
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "TeamSiteProvisioningWeb.Content.embed-script.js";

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                using (StreamReader reader = new StreamReader(stream))
                {
                    scriptBlock = reader.ReadToEnd();
                }

                var web = context.Web;
                var existingActions = web.UserCustomActions;
                context.Load(existingActions);
                context.ExecuteQuery();

                var actions = existingActions.ToArray();
                foreach (var action in actions)
                {
                    if (action.Description == "scenario1" && action.Location == "ScriptLink")
                    {
                        action.DeleteObject();
                        context.ExecuteQuery();
                    }
                }

                var newAction = existingActions.Add();
                newAction.Description = "scenario1";
                newAction.Location = "ScriptLink";

                newAction.ScriptBlock = scriptBlock;
                newAction.Update();
                context.Load(web, s => s.UserCustomActions);
                context.ExecuteQuery();
            }
        }
    }
}