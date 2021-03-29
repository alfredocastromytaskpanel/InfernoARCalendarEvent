/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph;
using InfernoARCalendarEvent.Extensions;
using InfernoARCalendarEvent.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace InfernoARCalendarEvent.Services
{
    public static class GraphService
    {
        private const string PlaceholderImage = "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==";

        // Load user's profile in formatted JSON.
        public static async Task<string> GetUserJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);

            try
            {
                // Load user profile.
                var user = await graphClient.Users[email].Request().GetAsync();
                return JsonConvert.SerializeObject(user, Formatting.Indented);
            }
            catch (ServiceException e)
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
                    default:
                        return JsonConvert.SerializeObject(new { Message = "An unknown error has occurred." }, Formatting.Indented);
                }
            }
        }

        // Load user's profile picture in base64 string.
        public static async Task<string> GetPictureBase64(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            try
            {
                // Load user's profile picture.
                //var pictureStream = await GetPictureStream(graphClient, email, httpContext);
                Stream pictureStream = null;

                if (pictureStream == null) return PlaceholderImage;

                // Copy stream to MemoryStream object so that it can be converted to byte array.
                var pictureMemoryStream = new MemoryStream();
                await pictureStream.CopyToAsync(pictureMemoryStream);

                // Convert stream to byte array.
                var pictureByteArray = pictureMemoryStream.ToArray();

                // Convert byte array to base64 string.
                var pictureBase64 = Convert.ToBase64String(pictureByteArray);

                return "data:image/jpeg;base64," + pictureBase64;
            }
            catch (Exception e)
            {
                return e.Message switch
                {
                    "ResourceNotFound" => PlaceholderImage, // If picture is not found, return the placeholder image.
                    "EmailIsNull" => JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented),
                    _ => null,
                };
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
        public static async Task<Stream> GetMyPictureStream(GraphServiceClient graphClient, HttpContext httpContext)
        {
            Stream pictureStream = null;

            try
            {
                try
                {
                    // Load user's profile picture.
                    pictureStream = await graphClient.Me.Photo.Content.Request().GetAsync();
                }
                catch (ServiceException e)
                {
                    if (e.Error.Code == "GetUserPhoto") // User is using MSA, we need to use beta endpoint
                    {
                        // Set Microsoft Graph endpoint to beta, to be able to get profile picture for MSAs 
                        graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                        // Get profile picture from Microsoft Graph
                        pictureStream = await graphClient.Me.Photo.Content.Request().GetAsync();

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

        // Send an email message from the current user.
        public static async Task SendEmail(GraphServiceClient graphClient, IWebHostEnvironment hostingEnvironment, string recipients, HttpContext httpContext)
        {
            if (recipients == null) return;

            var attachments = new MessageAttachmentsCollectionPage();

            try
            {
                // Load user's profile picture.
                //var pictureStream = await GetMyPictureStream(graphClient, httpContext);
                Stream pictureStream = null;

                if (pictureStream != null)
                {
                    // Copy stream to MemoryStream object so that it can be converted to byte array.
                    var pictureMemoryStream = new MemoryStream();
                    await pictureStream.CopyToAsync(pictureMemoryStream);

                    // Convert stream to byte array and add as attachment.
                    attachments.Add(new FileAttachment
                    {
                        ODataType = "#microsoft.graph.fileAttachment",
                        ContentBytes = pictureMemoryStream.ToArray(),
                        ContentType = "image/png",
                        Name = "me.png"
                    });
                }
            }
            catch (Exception e)
            {
                switch (e.Message)
                {
                    case "ResourceNotFound":
                        break;
                    default:
                        throw;
                }
            }

            // Prepare the recipient list.
            var splitRecipientsString = recipients.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            var recipientList = splitRecipientsString.Select(recipient => new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient.Trim()
                }
            }).ToList();

            // Build the email message.
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = System.IO.File.ReadAllText(hostingEnvironment.WebRootPath + "/email_template.html"),
                    ContentType = BodyType.Html,
                },
                Subject = "Sent from the Microsoft Graph Connect sample",
                ToRecipients = recipientList,
                Attachments = attachments
            };

            await graphClient.Me.SendMail(email, true).Request().PostAsync();
        }

        public static async Task<bool> CreateEvent(GraphServiceClient graphClient, string InfernoAPIKey, string recipients, string eventId, HttpContext httpContext)
        {
            if (recipients == null) return false;

            eventId = eventId ?? "248d8ea0-b518-493d-b9c1-0a9f3e4e94c7";

            try
            {
                var me = await graphClient.Me.Request().GetAsync();

                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + InfernoAPIKey);

                    Event newEvent = null;
                    string tzName = "Pacific Standard Time";
                    try
                    {
                        var response = await httpClient.GetAsync("https://api.infernocore.jolokia.com/api/Events/" + eventId);
                        if (!response.IsSuccessStatusCode) return false;
                        string apiResponse = await response.Content.ReadAsStringAsync();
                        var infEvent = JsonConvert.DeserializeObject<InfernoEvent>(apiResponse);
                        tzName = infEvent.startTime.GetTimeZoneStandardName();
                        newEvent = infEvent.ToMSGraphEvent();
                    }
                    catch (Exception)
                    {
                        //Default event just to keep debug without api.infernocore
                        newEvent = new Event
                        {
                            Subject = "Let's go for lunch",
                            Body = new ItemBody
                            {
                                ContentType = BodyType.Html,
                                Content = "Does noon work for you?"
                            },
                            Start = new DateTimeTimeZone
                            {
                                DateTime = "2021-03-30T10:00:00",
                                TimeZone = "Pacific Standard Time"
                            },
                            End = new DateTimeTimeZone
                            {
                                DateTime = "2021-03-30T11:00:00",
                                TimeZone = "Pacific Standard Time"
                            },
                            Attendees = new List<Attendee>()
                            {
                                new Attendee
                                {
                                    EmailAddress = new EmailAddress
                                    {
                                        Address = me.UserPrincipalName,
                                        Name = me.DisplayName
                                    },
                                    Type = AttendeeType.Required
                                }
                            }
                        };
                    }

                    var recipList = recipients.Split(";").ToList();
                    foreach (var recip in recipList)
                    {
                        //TODO Validate Email
                        try
                        {
                            var userAttend = await graphClient.Users[recip].Request().GetAsync();
                            newEvent.Attendees.ToList().Add(
                                    new Attendee
                                    {
                                        EmailAddress = new EmailAddress
                                        {
                                            Address = recip,
                                            Name = userAttend.DisplayName
                                        },
                                        Type = AttendeeType.Required
                                    });
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }

                    await graphClient.Me.Events
                            .Request()
                            .Header("Prefer", $"outlook.timezone=\"{tzName}\"") //"outlook.timezone=\"Pacific Standard Time\""
                            .AddAsync(newEvent);

                    return true;
                }
            }
            catch (Exception)
            {
                //TODO
            }

            return false;
        }
    }
}

/*var newEvent = new Event
{
    Subject = infEvent.name, //"Let's go for lunch"
    Body = new ItemBody
    {
        ContentType = BodyType.Html,
        Content = infEvent.name //"Does noon work for you?"
    },
    Start = new DateTimeTimeZone
    {
        DateTime = infEvent.startTime.DateTime.ToString("yyyy-MM-ddTHH:mm:ss"), //"yyyy'-'MM'-'dd'T'HH':'mm':'ss" "2021-03-28T12:00:00",
        TimeZone = tzName //"Pacific Standard Time"
    },
    End = new DateTimeTimeZone
    {
        DateTime = infEvent.startTime.DateTime.ToString("yyyy-MM-ddTHH:mm:ss"), //"2021-03-28T14:00:00",
        TimeZone = tzName //"Pacific Standard Time"
    },
    //Location = new Location
    //{
    //    DisplayName = "Harry's Bar"
    //},
    Attendees = new List<Attendee>()
    {
        new Attendee
        {
            EmailAddress = new EmailAddress
            {
                Address = recipients, 
                Name = "Alfredo Castro"
            },
            Type = AttendeeType.Required
        }
    },
    AllowNewTimeProposals = true,
};*/

