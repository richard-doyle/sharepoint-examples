using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TeamSiteProvisioningWeb.Helpers;
using TeamSiteProvisioningWeb.Models;

namespace TeamSiteProvisioningWeb.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            var details = new SiteDetails();
            return View();
        }

        [HttpPost]
        public ActionResult Index(SiteDetails details)
        {
            if (!string.IsNullOrWhiteSpace(details.Title))
            {
                var provisioner = new SiteProvisioner();
                provisioner.ProvisionSite(details);
            }

            return Redirect("/");
        }

        [HttpGet]
        public ActionResult Search()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Search(string searchTerm)
        {
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                var search = new SiteSearch();
                var results = search.Search(searchTerm);
                foreach (var item in results)
                {
                    System.Console.WriteLine(item);
                }
            }

            return Redirect("/Home/Search");
        }

        public ActionResult Details()
        {
            var siteInformationRetriever = new SiteInformationRetriever();
            var siteInfo = siteInformationRetriever.GetInfo();

            return View(siteInfo);
        }

        public ActionResult EmbedScript()
        {
            var scriptEmbedder = new ScriptEmbedder();
            scriptEmbedder.Embed();
            return View();
        }

        public ActionResult AddAppPart()
        {
            var appPartAdder = new AppPartAdder();
            appPartAdder.AddAppPart();
            return View();
        }
    }
}
