using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Kiota.Authentication.Azure;
using Serilog;
using System;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;

namespace ExchangeGraphTool
{
    /// <summary>
    /// Cloud instance
    /// </summary>
    public enum CloudInstance
    {
        /// <summary>
        /// Global cloud service
        /// </summary>
        Global,

        /// <summary>
        /// Microsoft Graph China operated by 21Vianet
        /// </summary>
        China,

        /// <summary>
        /// Microsoft Graph for US Government L4
        /// </summary>
        US_GOV,

        /// <summary>
        /// Microsoft Graph for US Government L5 (DOD)
        /// </summary>
        US_GOV_DOD
    }

    internal class GraphApiFactory : IDisposable
    {
        /// <summary>
        /// Graph API client instance
        /// </summary>
        public GraphServiceClient Client { get; private set; }

        /// <summary>
        /// Graph endpoint base URI
        /// </summary>
        public Uri GraphBaseUri
        {
            get
            {
                if (Client?.RequestAdapter?.BaseUrl == null)
                    throw new InvalidOperationException("Graph client not initialized");

                var uri = new Uri(Client.RequestAdapter.BaseUrl);
                return new Uri(uri.GetLeftPart(UriPartial.Authority));
            }
        }

        /// <summary>
        /// Authority host URI
        /// </summary>
        public Uri AuthorityHost { get; set; }

        /// <summary>
        /// Create an instance of Graph API using Azure AD application credentials
        /// </summary>
        /// <param name="clientId">Application/Client ID</param>
        /// <param name="tenantId">Tenant ID</param>
        /// <param name="clientSecret">Client secret if clientCert not provided</param>
        /// <param name="clientCert">Client certificate. If not provided, clientSecret must be provided</param>
        /// <param name="cloudInstance">Azure/Graph cloud instance</param>
        public GraphApiFactory(string clientId, string tenantId, string? clientSecret = null, X509Certificate2? clientCert = null, CloudInstance cloudInstance = CloudInstance.Global)
        {
            if (string.IsNullOrEmpty(clientId))
                throw new ArgumentNullException(nameof(clientId));

            if (string.IsNullOrEmpty(clientSecret) && clientCert == null)
                throw new ArgumentNullException(nameof(clientCert), "Must provide either clientCert or clientSecret");

            AuthorityHost = GetAzureCloudInstance(cloudInstance);
            string graphCloudInstance = GetGraphCloudInstance(cloudInstance);

            var azureClientOptions = new ClientSecretCredentialOptions { AuthorityHost = AuthorityHost };

            TokenCredential clientCredential;
            if (clientCert != null)
                clientCredential = new ClientCertificateCredential(tenantId, clientId, clientCert, azureClientOptions);
            else
                clientCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, azureClientOptions);

            // Set permissions scope based on the cloud instance
            var baseCloudUri = GetGraphCloudBaseUri(cloudInstance);
            string[] scopes = [$"{baseCloudUri.Scheme}://{baseCloudUri.Authority}/.default"];

            Log.Information("Creating Graph API client for cloud instance {@instance}", cloudInstance);
            Log.Information("Authority host: {host}", AuthorityHost);
            Log.Information("Graph cloud instance URL {baseUrl}", baseCloudUri);
            Log.Information("Scopes: {@scopes}", scopes);
            Log.Information("Using client ID {clientId} in tenant {tenantId}", clientId, tenantId);
            Log.Information("Authenticating with {authType}", clientCert != null ? "client certificate" : "client secret");

            var authProvider = new AzureIdentityAuthenticationProvider(clientCredential, scopes: scopes);
            
            string version = $"{Program.AppVersion?.Major}.{Program.AppVersion?.Minor}.{Program.AppVersion?.Build}";

            var httpClient = GraphClientFactory.Create(authProvider, nationalCloud: graphCloudInstance);
            httpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("ExchangeGraphTool", version));

            if (httpClient.BaseAddress == null)
                throw new InvalidOperationException("Graph client base address not initialized");

            Client = new GraphServiceClient(httpClient, baseUrl: httpClient.BaseAddress.ToString());
        }

        private static Uri GetAzureCloudInstance(CloudInstance cloudInstance)
        {
            return cloudInstance switch
            {
                CloudInstance.China => AzureAuthorityHosts.AzureChina,
                CloudInstance.US_GOV => AzureAuthorityHosts.AzureGovernment,
                CloudInstance.US_GOV_DOD => AzureAuthorityHosts.AzureGovernment,
                _ => AzureAuthorityHosts.AzurePublicCloud
            };
        }

        private static string GetGraphCloudInstance(CloudInstance cloudInstance)
        {
            return cloudInstance switch
            {
                CloudInstance.China => GraphClientFactory.China_Cloud,
                CloudInstance.US_GOV => GraphClientFactory.USGOV_Cloud,
                CloudInstance.US_GOV_DOD => GraphClientFactory.USGOV_DOD_Cloud,
                _ => GraphClientFactory.Global_Cloud,
            };
        }

        /// Microsoft Graph service national cloud endpoints
        /// From GraphClientFactory class in Microsoft Graph SDK. These are inaccessible in the SDK as they are private.
        /// If they change in the SDK, they need updating here
        private static readonly Dictionary<CloudInstance, Uri> _cloudList = new()
        {
            { CloudInstance.Global, new Uri("https://graph.microsoft.com") },
            { CloudInstance.US_GOV, new Uri("https://graph.microsoft.us") },
            { CloudInstance.China, new Uri("https://microsoftgraph.chinacloudapi.cn") },
            { CloudInstance.US_GOV_DOD, new Uri("https://dod-graph.microsoft.us") }
        };

        private static Uri GetGraphCloudBaseUri(CloudInstance cloudInstance)
        {
            return _cloudList.TryGetValue(cloudInstance, out var uri) ? uri : _cloudList[CloudInstance.Global];
        }

        #region IDisposable
        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // dispose managed state (managed objects)
                    Client.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~GraphApiFactory()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
