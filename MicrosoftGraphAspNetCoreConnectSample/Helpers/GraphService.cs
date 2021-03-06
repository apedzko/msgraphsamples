﻿/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public static class GraphService
    {
        // Load user's profile in formatted JSON.
        public static async Task<string> GetUserJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null)
                return EmailMissingError();

            try
            {
                // Load user profile.
                var user = await graphClient.Users[email].Request().GetAsync();
                return JsonConvert.SerializeObject(user, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return await FormatError(email, httpContext, e);
            }
        }

        // Load user's profile picture in base64 string.
        public static async Task<string> GetPictureBase64(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            try
            {
                // Load user's profile picture.
                using (var pictureStream = await GetPictureStream(graphClient, email, httpContext))
                {

                    // Copy stream to MemoryStream object so that it can be converted to byte array.
                    using (var pictureMemoryStream = new MemoryStream())
                    {
                        await pictureStream.CopyToAsync(pictureMemoryStream);

                        // Convert stream to byte array.
                        var pictureByteArray = pictureMemoryStream.ToArray();

                        // Convert byte array to base64 string.
                        var pictureBase64 = Convert.ToBase64String(pictureByteArray);

                        return "data:image/jpeg;base64," + pictureBase64;
                    }
                }
            }
            catch (Exception e)
            {
                switch (e.Message)
                {
                    case "ResourceNotFound":
                        // If picture not found, return the default image.
                        return "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==";
                    case "EmailIsNull":
                        return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);
                    default:
                        return null;
                }
            }
        }

        public static async Task<Stream> GetPictureStream(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) throw new Exception("EmailIsNull");

            Stream pictureStream = null;

            try
            {
                try
                {
                    // Load user's profile picture.
                    pictureStream = await graphClient.Users[email].Photo.Content.Request().GetAsync();
                }
                catch (ServiceException e)
                {
                    if (e.Error.Code == "GetUserPhoto") // User is using MSA, we need to use beta endpoint
                    {
                        // Set Microsoft Graph endpoint to beta, to be able to get profile picture for MSAs 
                        graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                        // Get profile picture from Microsoft Graph
                        pictureStream = await graphClient.Users[email].Photo.Content.Request().GetAsync();

                        // Reset Microsoft Graph endpoint to v1.0
                        graphClient.BaseUrl = "https://graph.microsoft.com/v1.0";
                    }
                }
            }
            catch (ServiceException e)
            {
                switch (e.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                    case "ErrorInvalidUser":
                        // If picture not found, return the default image.
                        throw new Exception("ResourceNotFound");
                    case "TokenNotFound":
                        await httpContext.ChallengeAsync();
                        return null;
                    default:
                        return null;
                }
            }

            return pictureStream;
        }


        public static async Task<string> GetUserCalendarJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);

            try
            {
                // Load user's calendar.
                var resultPage = await graphClient.Users[email].Events.Request()
                   // Only return the fields used by the application
                   //.Select(e => new {
                   //    e.Subject,
                   //    e.Organizer,
                   //    e.Attendees,
                   //    e.BodyPreview,
                   //    e.Start,
                   //    e.End
                   //})
                   // Sort results by when they were created, newest first
                   .OrderBy("createdDateTime DESC")
                   .GetAsync();

                return JsonConvert.SerializeObject(resultPage.CurrentPage, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return await FormatError(email, httpContext, e);
            }
        }

        public static async Task<string> GetUserInsightsJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);

            try
            {
                // Load user's calendar.
                var resultPage = await graphClient.Users[email].Insights.Used.Request()
                   .GetAsync();

                return JsonConvert.SerializeObject(resultPage, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return await FormatError(email, httpContext, e);
            }
        }


        public static async Task<string> GetUserFilesJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);

            try
            {
                // Load user's calendar.
                var resultPage = await graphClient.Users[email].Drive.Recent().Request()
                   .GetAsync();

                return JsonConvert.SerializeObject(resultPage, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return await FormatError(email, httpContext, e);
            }
        }

        public static async Task<string> GetUserReceivedEmailJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);

            try
            {
                // Load user's calendar.
                var resultPage = await graphClient.Users[email].MailFolders.Inbox.Messages.Request()
                   .GetAsync();

                return JsonConvert.SerializeObject(resultPage, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return await FormatError(email, httpContext, e);
            }
        }

        public static async Task<string> GetUserSentEmailJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null)
                return EmailMissingError();

            try
            {
                // Load user's calendar.
                var resultPage = await graphClient.Users[email].MailFolders.SentItems.Messages.Request()
                   .GetAsync();

                return JsonConvert.SerializeObject(resultPage, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return await FormatError(email, httpContext, e);
            }
        }

        public static async Task<List<ItemDetails>> GetAllItems(GraphServiceClient graphClient, string email)
        {
            var result = new List<ItemDetails>();

            if (email == null) return await Task.FromResult(result);

            foreach (var @event in await graphClient.Users[email].Events.Request().GetAsync())
            {
                result.Add(new ItemDetails { CreatedAt = @event.CreatedDateTime, Type = "Outlook Calendar Event", Name = @event.Subject });
            }

            foreach (var insight in await graphClient.Users[email].Insights.Used.Request().GetAsync())
            {
                result.Add(new ItemDetails { CreatedAt = insight.LastUsed.LastAccessedDateTime, Type = "Recently Used File (Insights)", Name = insight.ResourceVisualization.Title});
            }

            foreach (var file in await graphClient.Users[email].Drive.Recent().Request().GetAsync())
            {
                result.Add(new ItemDetails { CreatedAt = file.CreatedDateTime, Type = "Recently Used File", Name = file.Name });
            }

            foreach (var receivedMessage in await graphClient.Users[email].MailFolders.Inbox.Messages.Request().GetAsync())
            {
                result.Add(new ItemDetails { CreatedAt = receivedMessage.CreatedDateTime, Type = "Outlook Email Received", Name = receivedMessage.Subject });
            }

            foreach (var sentItem in await graphClient.Users[email].MailFolders.SentItems.Messages.Request().GetAsync())
            {
                result.Add(new ItemDetails { CreatedAt = sentItem.CreatedDateTime, Type = "Outlook Email Sent", Name = sentItem.Subject });
            }

            return await Task.FromResult(result);
        }

        private static string EmailMissingError()
        {
            return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);
        }

        private static async Task<string> FormatError(string email, HttpContext httpContext, ServiceException e)
        {
            switch (e.Error.Code)
            {
                case "Request_ResourceNotFound":
                case "ResourceNotFound":
                case "ErrorItemNotFound":
                case "itemNotFound":
                    return JsonConvert.SerializeObject(new { Message = $"User '{email}' was not found." }, Formatting.Indented);
                case "ErrorInvalidUser":
                    return JsonConvert.SerializeObject(new { Message = $"The requested user '{email}' is invalid." }, Formatting.Indented);
                case "AuthenticationFailure":
                    return JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                case "TokenNotFound":
                    await httpContext.ChallengeAsync();
                    return JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                case "ErrorAccessDenied":
                    return JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                default:
                    return JsonConvert.SerializeObject(new { Message = "An unknown error has occurred." }, Formatting.Indented);
            }
        }
    }
}
