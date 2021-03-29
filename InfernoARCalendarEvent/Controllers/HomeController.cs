/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.AspNetCore.Hosting;
using System.Security.Claims;
using InfernoARCalendarEvent.Services;
using Microsoft.Extensions.Configuration;
using System;

namespace InfernoARCalendarEvent.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly IGraphServiceClientFactory _graphServiceClientFactory;
        private readonly IConfiguration _configuration;

        public HomeController(IWebHostEnvironment hostingEnvironment, IGraphServiceClientFactory graphServiceClientFactory, IConfiguration configuration)
        {
            _env = hostingEnvironment;
            _graphServiceClientFactory = graphServiceClientFactory;
            _configuration = configuration;
        }

        [AllowAnonymous]
        // Load user's profile.
        public async Task<IActionResult> Index(string email)
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                email ??= User.FindFirst("preferred_username")?.Value;
                ViewData["Email"] = email;

                // Initialize the GraphServiceClient.
                var graphClient = _graphServiceClientFactory.GetAuthenticatedGraphClient((ClaimsIdentity)User.Identity);

                ViewData["Response"] = await GraphService.GetUserJson(graphClient, email, HttpContext);

                ViewData["Picture"] = await GraphService.GetPictureBase64(graphClient, email, HttpContext);
            }

            return View();
        }

        [Authorize]
        [HttpPost]
        // Send an email message from the current user.
        public async Task<IActionResult> SendEmail(string recipients)
        {
            if (string.IsNullOrEmpty(recipients))
            {
                TempData["Message"] = "Please add a valid email address to the recipients list!";
                return RedirectToAction("Index");
            }

            try
            {
                // Initialize the GraphServiceClient.
                var graphClient = _graphServiceClientFactory.GetAuthenticatedGraphClient((ClaimsIdentity)User.Identity);

                // Send the email.
                await GraphService.SendEmail(graphClient, _env, recipients, HttpContext);

                // Reset the current user's email address and the status to display when the page reloads.
                TempData["Message"] = "Success! Your mail was sent.";
                return RedirectToAction("Index");
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "Caller needs to authenticate.") return new EmptyResult();
                return RedirectToAction("Error", "Home", new { message = "Error: " + se.Error.Message });
            }
        }

        [Authorize]
        [HttpPost]
        // Create calendar event.
        public async Task<IActionResult> CreateEvent(string recipients, string eventId)
        {
            if (string.IsNullOrEmpty(recipients))
            {
                TempData["Message"] = "Please add a valid email address to the recipients list!";
                return RedirectToAction("Index");
            }

            try
            {
                string InfernoAPIKey = _configuration.GetValue<string>("InfernoAPIKey");

                // Initialize the GraphServiceClient.
                var graphClient = _graphServiceClientFactory.GetAuthenticatedGraphClient((ClaimsIdentity)User.Identity);

                // Send the email.
                var result = await GraphService.CreateEvent(graphClient, InfernoAPIKey, recipients, eventId, HttpContext);

                // Reset the current user's email address and the status to display when the page reloads.
                TempData["Message"] = result ? "Success! Your calendar event was created." : "";
                return RedirectToAction("Index");
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "Caller needs to authenticate.") return new EmptyResult();
                return RedirectToAction("Error", "Home", new { message = "Error: " + se.Error.Message });
            }
            catch (Exception e)
            {
                return RedirectToAction("Error", "Home", new { message = "Error: " + e.Message });
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [AllowAnonymous]
        public IActionResult Error()
        {
            return View();
        }
    }
}
