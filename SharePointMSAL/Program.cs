using System;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;

namespace SharePointMSAL
{
    class Program
    {
        static async Task Main(string[] args)
        {
            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", true, true)
            .Build();

            string siteUrl = $"https://{config["environment"]}.sharepoint.com{config["site"]}";
            string identity = $"{config["environment"]}-{config["name"]}";           
            string keyVaultName = GetKeyVaultName(identity);
            string certificateName = identity;
            string tenantId = $"{config["environment"]}.onmicrosoft.com";

            string clientIDSecret = identity.Replace("_","").Replace("-","");
            string clientId = GetSecretFromKeyVault(keyVaultName, clientIDSecret);

            //For SharePoint app only auth, the scope will be the Sharepoint tenant name followed by /.default
            var scopes = new string[] { $"https://{config["environment"]}.sharepoint.com/.default" };

            var accessToken = await GetApplicationAuthenticatedClient(clientId, keyVaultName, certificateName, scopes, tenantId);

            var ctx = GetClientContextWithAccessToken(siteUrl, accessToken);

            Web web = ctx.Web;
            ctx.Load(web);
            await ctx.ExecuteQueryAsync();

            Console.WriteLine(web.Title);
        }

        private static string GetKeyVaultName(string identity)
        {
            var keyVaultName = identity;

            if (keyVaultName.Length > 24)
            {
                keyVaultName = keyVaultName.Substring(0, 24);
            }
            return keyVaultName;
        }
        private static async Task<string> GetApplicationAuthenticatedClient(string clientId, string keyVaultName, string certificateName, string[] scopes, string tenantId)
        {
            var certificate = GetAppOnlyCertificate(keyVaultName, certificateName);
            IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder
                                                        .Create(clientId)
                                                        .WithCertificate(certificate)
                                                        .WithTenantId(tenantId)
                                                        .Build();

            AuthenticationResult authResult = await clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
            string accessToken = authResult.AccessToken;
            return accessToken;
        }

        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);
            clientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
            {
                webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }

        public static X509Certificate2 GetAppOnlyCertificate(string keyVaultName, string certificateName)
        {
            var keyVaultUrl = $"https://{keyVaultName}.vault.azure.net";

            var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());
            KeyVaultSecret keyVaultSecret = client.GetSecret(certificateName);

            X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String(keyVaultSecret.Value), string.Empty,
              X509KeyStorageFlags.MachineKeySet |
              X509KeyStorageFlags.PersistKeySet |
              X509KeyStorageFlags.Exportable);

            return certificate;
        }

        public static string GetSecretFromKeyVault(string keyVaultName, string secretName){
            var keyVaultUrl = $"https://{keyVaultName}.vault.azure.net";

            var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());
            KeyVaultSecret keyVaultSecret = client.GetSecret(secretName);

            return keyVaultSecret.Value;
        }
    }
}

