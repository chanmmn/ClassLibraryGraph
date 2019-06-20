using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ClassLibraryGraph
{
    public class MSGraph
    {
        public static List<string> lstString = new List<string>();
        public static async Task<List<string>> CallRunAsync()
        {

            await Run();
            return lstString;
        }

        #region POC Code

        //    public static async Task Run()
        //    {
        //        var clientId = "7515c9fd-d9d2-4aa5-b960-d5b3c0032502";
        //        var scopes = new List<string>() { "User.ReadBasic.All" }.ToArray();

        //        IPublicClientApplication clientApplication = InteractiveAuthenticationProvider.CreateClientApplication(clientId);
        //        InteractiveAuthenticationProvider authProvider = new InteractiveAuthenticationProvider(clientApplication, scopes);

        //        GraphServiceClient graphClient = new GraphServiceClient(authProvider);

        //        //var users = await graphClient.Users
        //        //    .Request()
        //        //    .Select(e => new {
        //        //        e.DisplayName,
        //        //        e.GivenName,
        //        //        e.PostalCode,
        //        //        e.Manager
        //        //    })
        //        //    .GetAsync();
        //        var users = await graphClient.Users
        //.Request()
        //.Select(e => new
        //{
        //    e.DisplayName,
        //    e.GivenName,
        //    e.PostalCode,
        //    e.Manager
        //})
        //.GetAsync();

        //        foreach (var user in users)
        //        {
        //            //Console.WriteLine(user.GivenName);
        //            lstString.Add(user.DisplayName);
        //        }
        //    }

        #endregion

        public static async Task Run()
        {
            var clientId = "Cliebt ID";
            var scopes = new List<string>() { "User.ReadBasic.All" }.ToArray();

            IPublicClientApplication clientApplication = InteractiveAuthenticationProvider.CreateClientApplication(clientId);
            InteractiveAuthenticationProvider authProvider = new InteractiveAuthenticationProvider(clientApplication, scopes);

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);


            var users = await graphClient.Users
                .Request()
                .Select(e => new
                {
                    e.DisplayName,
                    e.GivenName,
                    e.PostalCode,
                    e.Manager
                })
                .GetAsync();

            foreach (var user in users)
            {
                //Console.WriteLine(user.GivenName);
                lstString.Add(user.DisplayName);
            }

        }
    }
}
