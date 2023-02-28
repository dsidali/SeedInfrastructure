using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace SeedInfrastructure.MicrosoftTeams
{
    internal class MsalPublicClientAuthenticationProvider : IAuthenticationProvider
    {
        private static MsalPublicClientAuthenticationProvider _singleton;
        private IPublicClientApplication _application;
        private string[] _scopes;

        private MsalPublicClientAuthenticationProvider(IPublicClientApplication application, string[] scopes)
        {
            _application = application;
            _scopes = scopes;
        }

        public static MsalPublicClientAuthenticationProvider GetInstance(IPublicClientApplication application, string[] scopes)
        {
            if (_singleton == null)
            {
                _singleton = new MsalPublicClientAuthenticationProvider(application, scopes);
            }

            return _singleton;
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", await GetTokenAsync());
        }

        public async Task<string> GetTokenAsync()
        {
            AuthenticationResult result = null;

            try
            {
                result = await _application.AcquireTokenInteractive(_scopes).ExecuteAsync();
                return result.AccessToken;
            }
            catch (MsalServiceException e)
            {
                Console.WriteLine(e.Claims);
                return result.AccessToken;
            }


        }
    }
}
