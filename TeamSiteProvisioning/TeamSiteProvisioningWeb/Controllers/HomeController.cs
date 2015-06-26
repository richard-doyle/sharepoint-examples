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
    }
}
