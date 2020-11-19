using System;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using GraphConsoleApp.Controller;

namespace GraphConsoleApp
{
  class Program
  {
    static void Main(string[] args)
    {      
      var config = LoadAppSettings();
      if (config != null)
      {
        string authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";
        GraphController controller = new GraphController();
        controller.Initialize(config["clientId"], authority, "", config["clientSecret"]);

        string[] addresses = GetMailAddresses();
        foreach (string a in addresses)
        {
          var userID = controller.GetUserId(a).Result;
          Console.WriteLine("Address: " + a);
          Console.WriteLine("ID: " + userID);
        }

        string[] batchedUserIDs = controller.GetUserIds(addresses).Result;

        foreach (string id in batchedUserIDs)
        {
          var stream = controller.GetUserPhoto(id).Result;
          if (stream != null)
          {
            var photo = controller.uploadUserPotoOneDrive(stream, $"{id}.jpg").Result;
            Console.WriteLine("Added file to: " + photo.WebUrl);
          }
        }
      }
    }

    static async void graphExec(IConfigurationRoot config)
    {
      
    }
    static string[] GetMailAddresses()
    {
      using (StreamReader file = File.OpenText(@"C:\temp\TestUsers.json"))
      {
        JsonSerializer serializer = new JsonSerializer();
        string[] movie2 = (string[])serializer.Deserialize(file, typeof(string[]));
        return movie2;
      }
    }

    private static IConfigurationRoot LoadAppSettings()
    {
      try
      {
        string currentPath = System.IO.Directory.GetCurrentDirectory();
        var config = new ConfigurationBuilder()
                        .SetBasePath(currentPath)
                        .AddJsonFile("appsettings.json", false, true)
                        .Build();

        if (string.IsNullOrEmpty(config["clientId"]) ||
            string.IsNullOrEmpty(config["clientSecret"]) ||            
            string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }

        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }
  }
}
