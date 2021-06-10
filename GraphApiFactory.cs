using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Threading.Tasks;

namespace ExchangeGraphTool
{
    /// <summary>
    /// Cloud instance
    /// </summary>
    [JsonConverter(typeof(StringEnumConverter))]
    public enum CloudInstance
    {
        Global,
        China,
        Germany,
        US_GOV
    }

    internal class GraphApiFactory
    {
        public GraphServiceClient Client { get; private set; }

        public string GetGraphBaseUrl()
        {
            var uri = new Uri(Client.BaseUrl);
            return uri.GetLeftPart(UriPartial.Authority);
        }


        public async Task<string> GetAzureADBaseUrl()
        {
            var uri = await _confidentialClientApplication.GetAuthorizationRequestUrl(new[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            return uri.GetLeftPart(UriPartial.Authority);
        }

        private readonly IConfidentialClientApplication _confidentialClientApplication;

        public GraphApiFactory(string clientId, string tenantId, string clientSecret, CloudInstance cloudInstance = CloudInstance.Global)
        {
            var azureCloudInstance = GetAzureCloudInstance(cloudInstance);
            string graphCloudInstance = GetGraphCloudInstance(cloudInstance);

            _confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(azureCloudInstance, tenantId)
                .WithClientSecret(clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(_confidentialClientApplication);

            var httpClient = GraphClientFactory.Create(authProvider, nationalCloud: graphCloudInstance);

            // TODO: update if Graph SDK is updated to passing in httpClient or cloud instance without needing httpClient
            // (proposed fix still needs URL as param
            // https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/916
            // https://github.com/microsoftgraph/msgraph-sdk-dotnet/pull/959 )
            Client = new GraphServiceClient(httpClient.BaseAddress.ToString(), authProvider);
        }

        private static AzureCloudInstance GetAzureCloudInstance(CloudInstance cloudInstance)
        {
            return cloudInstance switch
            {
                CloudInstance.China => AzureCloudInstance.AzureChina,
                CloudInstance.Germany => AzureCloudInstance.AzureGermany,
                CloudInstance.US_GOV => AzureCloudInstance.AzureUsGovernment,
                _ => AzureCloudInstance.AzurePublic,
            };
        }

        private static string GetGraphCloudInstance(CloudInstance cloudInstance)
        {
            return cloudInstance switch
            {
                CloudInstance.China => GraphClientFactory.China_Cloud,
                CloudInstance.Germany => GraphClientFactory.Germany_Cloud,
                CloudInstance.US_GOV => GraphClientFactory.USGOV_Cloud,
                _ => GraphClientFactory.Global_Cloud,
            };
        }
    }
}
