using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace GraphConsoleApp.Controller
{
  class GraphController
  {
    private GraphServiceClient graphClient;

    public void Initialize(string clientId, string authority, string redirectUri, string clientSecret)
    {
      var clientApplication = ConfidentialClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .WithClientSecret(clientSecret)
                                              .Build();
      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");
      string accessToken = clientApplication.AcquireTokenForClient(scopes).ExecuteAsync().Result.AccessToken;
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                      requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                    }));
      this.graphClient = graphClient;
    }

    public async Task<string> GetUserId(string mailAddress)
    {
      User user = await this.graphClient.Users[mailAddress].Request().GetAsync();
      return user.Id;
    }

    public async Task<string[]> GetUserIds(string[] mailAddresses)
    {
      var batchRequestContent = new BatchRequestContent();
      List<string> requestIDs = new List<string>();
      foreach (string a in mailAddresses)
      {
        var singleRequest = this.graphClient.Users[a].Request();
        string reqID = batchRequestContent.AddBatchRequestStep(singleRequest);
        requestIDs.Add(reqID);
      }
      var returnedResponse = await graphClient.Batch.Request().PostAsync(batchRequestContent);
      List<string> userIDs = new List<string>();  
      foreach (string id in requestIDs)
      {
        User u = await returnedResponse.GetResponseByIdAsync<User>(id);
        userIDs.Add(u.Id);
      }
      return userIDs.ToArray();
    }

    public async Task<Stream> GetUserPhoto(string userId)
    {
      try
      {
        var stream = await this.graphClient.Users[userId].Photo.Content
          .Request()
          .GetAsync();
        
        return stream;
      }
      catch(Exception ex)
      {
        if (ex.Message.StartsWith("Code: ErrorItemNotFound") || 
          (ex.InnerException != null && ex.InnerException.Message.StartsWith("Code: ErrorItemNotFound")))
        {
          Console.WriteLine("No user photo");
        }
        else
        {
          Console.WriteLine(ex);
        }
        return null;
      }
    }

    public async Task<DriveItem> uploadUserPotoOneDrive(Stream stream, string filename)
    {
      DriveItem uploadResult = await this.graphClient.Users["05de1e95-e588-464c-af66-fc4821d0b9c8"]
                                                    .Drive.Root
                                                    .ItemWithPath(filename)
                                                    .Content.Request()
                                                    .PutAsync<DriveItem>(stream);
      return uploadResult;
    }
  }
}
