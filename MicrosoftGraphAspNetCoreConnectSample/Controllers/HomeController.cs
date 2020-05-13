﻿/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MicrosoftGraphAspNetCoreConnectSample.Helpers;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.AspNetCore.Hosting;
using System.Security.Claims;

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IHostingEnvironment _env;
        private readonly IGraphSdkHelper _graphSdkHelper;

        public HomeController(IConfiguration configuration, IHostingEnvironment hostingEnvironment, IGraphSdkHelper graphSdkHelper)
        {
            _configuration = configuration;
            _env = hostingEnvironment;
            _graphSdkHelper = graphSdkHelper;
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

            ViewData["Email"] = "Recent Insights";

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

            ViewData["Email"] = "Recent Files";

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

                ViewData["Response"] = await GraphService.GetUserSentEmailJson(GetClient(), email, HttpContext);
            }

            ViewData["Title"] = "Sent Email";

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
