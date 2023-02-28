using Microsoft.Graph;
using Microsoft.Identity.Client;

using System.Collections.Generic;

namespace SeedInfrastructure.MicrosoftTeams
{
    public static class GraphAuthenticationHandler
    {
        private static IAuthenticationProvider CreateAuthorizationProviderConfidential(string tenantId, string applicationID, string clientSecret)
        {

            var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(applicationID)
                .WithAuthority(authority)
                .WithClientSecret(clientSecret)
                .Build();

            return MsalConfidentialClientAuthenticationProvider.GetInstance(cca, scopes.ToArray());

        }



        private static IAuthenticationProvider CreateAuthorizationProviderPublic(string tenantId, string applicationID, string clientSecret)
        {

            var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = PublicClientApplicationBuilder.Create(applicationID)
                .WithAuthority(authority)

                .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                .Build();

            return MsalPublicClientAuthenticationProvider.GetInstance(cca, scopes.ToArray());
        }


        public static GraphServiceClient GetAuthenticatedGraphClientConfidential(string tenantId, string applicationID, string clientSecret)
        {
            var authenticationProvider = CreateAuthorizationProviderConfidential(tenantId, applicationID, clientSecret);
            return new GraphServiceClient(authenticationProvider);


        }

        public static GraphServiceClient GetAuthenticatedGraphClientPublic(string tenantId, string applicationID, string clientSecret)
        {
            var authenticationProvider = CreateAuthorizationProviderPublic(tenantId, applicationID, clientSecret);
            return new GraphServiceClient(authenticationProvider);
        }
    }
}
