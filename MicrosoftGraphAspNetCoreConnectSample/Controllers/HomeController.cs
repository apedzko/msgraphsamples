/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MicrosoftGraphAspNetCoreConnectSample.Helpers;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    public class HomeController : Controller
    {
        private readonly IGraphSdkHelper _graphSdkHelper;
        private readonly IIManageService _iManageService;

        public HomeController(IGraphSdkHelper graphSdkHelper, IIManageService iManageService)
        {
            _graphSdkHelper = graphSdkHelper;
            _iManageService = iManageService;
        }

        [AllowAnonymous]
        // Load user's profile.
        public async Task<IActionResult> Index(string email)
        {
            if (User.Identity.IsAuthenticated)
            {
                email = GetEmail(email);
                ViewData["Email"] = email;

                var client = GetClient();

                ViewData["Response"] = await GraphService.GetUserJson(client, email, HttpContext);

                ViewData["Picture"] = await GraphService.GetPictureBase64(client, email, HttpContext);
            }

            return View();
        }

        [AllowAnonymous]
        // Load user's timeline
        public async Task<IActionResult> Timeline(string email)
        {
            ViewData["Title"] = "Timeline";

            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                email = GetEmail(email);
                ViewData["Email"] = email;

                var graphItems = await GraphService.GetAllItems(GetClient(), email);
                var iManageItems = await _iManageService.GetRecentDocumentsAsync(email);

                var model = graphItems.Union(iManageItems).OrderByDescending(item => item.CreatedAt);

                return View(model);
            }

            return View();
        }

        [AllowAnonymous]
        // Load user's calendar.
        public async Task<IActionResult> Calendar(string email)
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                email = GetEmail(email);
                ViewData["Email"] = email;

                ViewData["Response"] = await GraphService.GetUserCalendarJson(GetClient(), email, HttpContext);
            }

            return View();
        }

        [AllowAnonymous]
        // Load user's insights.
        public async Task<IActionResult> Insights()
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                string email = GetEmail(null);
                ViewData["Email"] = email;

                ViewData["Response"] = await GraphService.GetUserInsightsJson(GetClient(), email, HttpContext);
            }

            ViewData["Title"] = "Recent Insights";

            return View("Results");
        }

        [AllowAnonymous]
        // Load user's files.
        public async Task<IActionResult> Files()
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                string email = GetEmail(null);
                ViewData["Email"] = email;

                ViewData["Response"] = await GraphService.GetUserFilesJson(GetClient(), email, HttpContext);
            }

            ViewData["Title"] = "Recent Files";

            return View("Results");
        }

        [AllowAnonymous]
        // Load user's inbox.
        public async Task<IActionResult> ReceivedMail()
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                string email = GetEmail(null);
                ViewData["Email"] = email;

                ViewData["Response"] = await GraphService.GetUserReceivedEmailJson(GetClient(), email, HttpContext);
            }

            ViewData["Title"] = "Received Email";

            return View("Results");
        }

        [AllowAnonymous]
        // Load user's sent mail.
        public async Task<IActionResult> SentMail()
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                string email = GetEmail(null);
                ViewData["Email"] = email;

                ViewData["Response"] = await GraphService.GetUserSentEmailJson(GetClient(), email, HttpContext);
            }

            ViewData["Title"] = "Sent Email";

            return View("Results");
        }

        [AllowAnonymous]
        // Load user's recent documents
        public async Task<IActionResult> RecentDocuments()
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                string email = GetEmail(null);
                ViewData["Email"] = email;

                ViewData["Response"] = await _iManageService.GetRecentDocumentsJsonAsync(email);
            }

            ViewData["Title"] = "Recent Documents (iManage)";

            return View("Results");
        }

        private GraphServiceClient GetClient()
        {

            // Initialize the GraphServiceClient.
            return _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);
        }

        private string GetEmail(string email)
        {
            // Get users's email.
            email = email ?? User.FindFirst("preferred_username")?.Value;
            return email;
        }

        [AllowAnonymous]
        public IActionResult Error()
        {
            return View();
        }
    }
}
