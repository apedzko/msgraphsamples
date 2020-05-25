using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public class IManageService : IIManageService
    {
        private readonly string _baseUrl;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _grantType;
        private readonly string _scope;
        private readonly string _userPassword;

        public IManageService(IConfiguration configuration)
        {
            _baseUrl = configuration["iManage:BaseUrl"];
            _clientId = configuration["iManage:ClientId"];
            _clientSecret = configuration["iManage:ClientSecret"];
            _grantType = configuration["iManage:GrantType"];
            _scope = configuration["iManage:Scope"];
            _userPassword = configuration["iManage:UserPassword"];
        }

        public async Task<List<ItemDetails>> GetRecentDocumentsAsync(string email)
        {
            var recentDocumentsJson = await GetRecentDocumentsJsonAsync(email);

            var items = JObject.Parse(recentDocumentsJson)["data"]["results"]
                .AsJEnumerable()
                .Select(item => new ItemDetails
                {
                    CreatedAt = item["create_date"].Value<DateTime?>(),
                    Name = item["name"].Value<string>(),
                    Type = "Document (iManage)"
                })
                .ToList();

            return await Task.FromResult(items);
        }

        public async Task<string> GetRecentDocumentsJsonAsync(string email)
        {
            using (var handler = new HttpClientHandler())
            {
                ConfigureHttpClientHandler(handler);

                using (var client = new HttpClient(handler))
                {
                    (var accessToken, var customerID) = await GetAccessTokenAndCustomerIDAsync(email);

                    ConfigureXAuthTokenHeader(client, accessToken);

                    var response = await client.GetAsync(_baseUrl + "/work/api/v2/customers/" + customerID + "/recent-documents?activity=all");

                    if (!response.IsSuccessStatusCode)
                        throw new Exception($"Status code '{response.StatusCode}' has been received during getting iManage recent documents");

                    return await response.Content.ReadAsStringAsync();
                }
            }
        }

        /// <summary>
        /// Returns access token and customer ID
        /// </summary>
        /// <returns></returns>
        private async Task<(string, string)> GetAccessTokenAndCustomerIDAsync(string email)
        {
            using (var handler = new HttpClientHandler())
            {
                ConfigureHttpClientHandler(handler);

                using (var client = new HttpClient(handler))
                {
                    var responseToken = await client.PostAsync(_baseUrl + "/auth/oauth2/token", GetTokenRequestContent(email));

                    if (!responseToken.IsSuccessStatusCode)
                        throw new Exception($"Status code '{responseToken.StatusCode}' has been received during getting iManage access token");

                    var responseTokenContent = await responseToken.Content.ReadAsStringAsync();

                    var accessToken = JObject.Parse(responseTokenContent)["access_token"].Value<string>();

                    ConfigureXAuthTokenHeader(client, accessToken);

                    var responseUserInfo = await client.GetAsync(_baseUrl + "/api");

                    if (!responseUserInfo.IsSuccessStatusCode)
                        throw new Exception($"Status code '{responseUserInfo.StatusCode}' has been received during getting iManage user info");

                    var responseUserInfoContent = await responseUserInfo.Content.ReadAsStringAsync();

                    var customerID = JObject.Parse(responseUserInfoContent)["data"]["user"]["customer_id"].Value<string>();

                    return (accessToken, customerID);
                }
            }
        }

        /// <summary>
        /// Returns token request content
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        private FormUrlEncodedContent GetTokenRequestContent(string email)
        {
            var userName = GetUserName(email);

            var payload = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>("username", userName),
                new KeyValuePair<string, string>("password", _userPassword),
                new KeyValuePair<string, string>("grant_type", _grantType),
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", _scope)
            };

            return new FormUrlEncodedContent(payload);
        }

        /// <summary>
        /// Configures an object of <see cref="HttpClientHandler"/> to always ignore invalid SSL certificate
        /// </summary>
        private static void ConfigureHttpClientHandler(HttpClientHandler httpClientHandler)
        {
            httpClientHandler.ServerCertificateCustomValidationCallback
                = (request, certificate, chain, sslPolicyErrors) => true;
        }

        /// <summary>
        /// Configures a X-Auth-Token header of an object of <seealso cref="HttpClient"/>
        /// </summary>
        private static void ConfigureXAuthTokenHeader(HttpClient httpClient, string accessToken)
        {
            httpClient.DefaultRequestHeaders.Add("X-Auth-Token", accessToken);
        }

        /// <summary>
        /// Extracts a user name from his/her email
        /// </summary>
        private static string GetUserName(string email)
        {
            return email.Substring(0, email.IndexOf('@'));
        }
    }
}
