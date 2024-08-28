// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Graph.Me.Messages.Item.Attachments.CreateUploadSession;
using Microsoft.Graph.Models.Security;
using Microsoft.Graph.Users.Item.SendMail;
using System.IO.Pipes;
class GraphHelper
{
    // <UserAuthConfigSnippet>
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        var options = new DeviceCodeCredentialOptions
        {
            ClientId = settings.ClientId,
            TenantId = settings.TenantId,
            DeviceCodeCallback = deviceCodePrompt,
        };

        _deviceCodeCredential = new DeviceCodeCredential(options);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    // </UserAuthConfigSnippet>

    // <GetUserTokenSnippet>
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }
    // </GetUserTokenSnippet>

    // <GetUserSnippet>
    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
        });
    }
    // </GetUserSnippet>

    // <GetInboxSnippet>
    public static Task<MessageCollectionResponse?> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages
            .GetAsync((config) =>
            {
                // Only request specific properties
                config.QueryParameters.Select = new[] { "from", "isRead", "receivedDateTime", "subject" };
                // Get at most 25 results
                config.QueryParameters.Top = 25;
                // Sort by received time, newest first
                config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };
            });
    }
    // </GetInboxSnippet>

    // <SendMailSnippet>
    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Create a new message
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = "jamie.fraser.1968@outlook.com"
                    }
                }
            }
        };
        // Send the message
        await _userClient.Me
            .SendMail
            .PostAsync(new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody
            {
                Message = message
            });
    }
    // </SendMailSnippet>

#pragma warning disable CS1998
    // <MakeGraphCallSnippet>
    // This function serves as a playground for testing Graph snippets
    // or other code
    public async static Task MakeGraphCallAsync()
    {
        // INSERT YOUR CODE HERE
    }
    // </MakeGraphCallSnippet>
    public static async Task SendMailWithLargeAttachmentAsync(string subject, string body, string recipient, string fileNameWithPath)
    {
        var message = await CreateDraftMessage(recipient, subject, body);
        var uploadSession = await CreateUploadSessionToAttachLargeFileToMessageAsync(message.Id, fileNameWithPath);
        var uploadResult = await UploadAttachmentAsync(uploadSession, fileNameWithPath);
        await _userClient.Me.Messages[message.Id].Send.PostAsync();
        
        //if(uploadResult.UploadSucceeded)
        //{
        //    var msg = new Message
        //    {
        //        Body = new ItemBody
        //        {
        //            Content = body,
        //            ContentType = BodyType.Textdrfa
        //        },
        //        Subject = subject,
        //        ToRecipients = new List<Recipient>
        //        {
        //            new Recipient
        //            {
        //                EmailAddress = new EmailAddress
        //                {
        //                    Address = recipient
        //                }
        //            }
        //        }
        //    }
        //}
    }
    private static async Task<Message> CreateDraftMessage(string recipientAddress, string subject, string body)
    {
        var requestBody = new Message
        {
            Subject = subject,
            Importance = Importance.Low,
            Body = new ItemBody
            {
                ContentType = BodyType.Text,
                Content = body,
            },
            ToRecipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipientAddress,
                    },
                },
            }
        };

        // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
         
        var result = await _userClient.Me.Messages.PostAsync(requestBody);

        return result;
    }
    private static async Task<UploadSession>CreateUploadSessionToAttachLargeFileToMessageAsync(string messageId, string filePath)
    {
        var fileSize = File.OpenRead(filePath).Length;
        var largeAttachment = new AttachmentItem
        {
            AttachmentType = AttachmentType.File,
            Name = Path.GetFileName(filePath),
            Size = fileSize,
        };
        var uploadSessionRequestBody = new CreateUploadSessionPostRequestBody
        {
            AttachmentItem = largeAttachment,
        };
        var uploadSession = await _userClient.Me
                                             .Messages[messageId]
                                             .Attachments
                                             .CreateUploadSession
                                             .PostAsync(uploadSessionRequestBody);


        
        return uploadSession;
    }
    private static async Task<UploadResult<AttachmentItem>> UploadAttachmentAsync(UploadSession session, string filePath)
    {
        int maxChunkSize = 320 * 1024; // 320 KB
        var stream = System.IO.File.OpenRead(filePath);
        var fileUploadTask = new LargeFileUploadTask<AttachmentItem>(session, stream, maxChunkSize);
        var result = await fileUploadTask.UploadAsync();
        return result;
    }
}
