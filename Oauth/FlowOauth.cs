using System.Collections.Generic;
using System.IO;
using System;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;
using System.Linq;
using GoogleApis.Oauth.Interfaces;
using System.Runtime.InteropServices;

namespace GoogleApis.Oauth
{
    [ComVisible(true)]
    [Guid("1f4e2b79-8d5a-4c6b-89a1-3e2d7f56c1b0")]
    [ClassInterface(ClassInterfaceType.None)]
    public class FlowOauth: IFlowOauth
    {
        private const string OAUTH2_END_POINT = "https://accounts.google.com/o/oauth2/v2/auth";
        private const string TOKEN_END_POINT = "https://oauth2.googleapis.com/token";
        private const string REVOKE_TOKEN_END_POINT = "https://oauth2.googleapis.com/revoke";

        private OAuthCredentials credentials;

        public string AccessToken => credentials.AccessToken;
        public string ApiKey => credentials.ApiKey;

        public string Navigator { get; set; } = "chrome.exe";
        private int Status { get; set; }

        public FlowOauth()
        {
            credentials = new OAuthCredentials();
        }
        public void InitializeFlow(string credentialsClientPath, string credentialsTokenPath, string credentialsApiKeyPath, string scopes)
        {
            try
            {

                ReadCredentialsApiKey(credentialsApiKeyPath);
                ReadCredentialsClientPath(credentialsClientPath);

                if (File.Exists(credentialsTokenPath))
                {
                    ReadCredentialsTokenPath(credentialsTokenPath);

                    if (!IsValidToken())
                    {
                        string content = UpdateTokenAccess();

                        if (Status == 200)
                        {
                            if (SaveToken(content, credentialsTokenPath))
                            {
                                ReadCredentialsTokenPath(credentialsTokenPath);
                            }
                        }
                    }
                }
                else
                {
                    string content = RequestToken(scopes);

                    if (!string.IsNullOrEmpty(content))
                    {
                        if (SaveToken(content, credentialsTokenPath))
                        {
                            ReadCredentialsTokenPath(credentialsTokenPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error InitializeFlow : \n", ex);
            }
        }
        public bool RevokeToken()
        {
            try
            {
                if (!string.IsNullOrEmpty(credentials.AccessToken))
                {
                    var body = new FormUrlEncodedContent(new[]
                {
                        new KeyValuePair<string, string>("token", credentials.AccessToken)
                    });

                    var headers = new Dictionary<string, string>
                    {
                        { "Content-type", "application/x-www-form-urlencoded" }
                    };

                    var response = SendRequest(HttpMethod.Post, REVOKE_TOKEN_END_POINT, body, headers);

                    return Status == 200;
                }
                else
                {
                    throw new Exception("First starting method InitializaFlow");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error to revoke Token ", ex);
            }
        }
        private void ReadCredentialsClientPath(string credentialsClientPath)
        {
            try
            {

                if (File.Exists(credentialsClientPath))
                {
                    string clientContent = File.ReadAllText(credentialsClientPath);
                    var clientData = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(clientContent);

                    if (clientData == null)
                    {
                        throw new InvalidOperationException("Error reading client file");
                    }

                    var appType = clientData.ContainsKey("web") ? "web" : "installed";

                    credentials.ClientId = clientData[appType]["client_id"];
                    credentials.ClientSecret = clientData[appType]["client_secret"];
                    credentials.RedirectUri = clientData[appType]["redirect_uris"][0];
                }
                else
                {
                    throw new FileNotFoundException($"File not found {credentialsClientPath}");
                }
            }
            catch (FileNotFoundException ex)
            {
                throw new InvalidOperationException($"Error reading client file: {credentialsClientPath}. File not found.", ex);
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Error parsing JSON from client file: {credentialsClientPath}", ex);
            }
            catch (KeyNotFoundException ex)
            {
                throw new InvalidOperationException($"Required key not found in client file: {credentialsClientPath}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unexpected error reading client file: {credentialsClientPath}", ex);
            }
        }
        private void ReadCredentialsTokenPath(string credentialsTokenPath)
        {
            try
            {
                var content = File.ReadAllText(credentialsTokenPath);
                var tokenData = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(content);

                if (tokenData == null)
                {
                    throw new InvalidOperationException("Error reading client file");
                }

                if (tokenData.ContainsKey("error") || tokenData.ContainsKey("error_description"))
                {
                    throw new InvalidOperationException($"Error reading token file: {tokenData["error"]}");
                }
                else if (tokenData.ContainsKey("access_token"))
                {
                    credentials.AccessToken = tokenData["access_token"];
                    credentials.RefreshToken = tokenData["refresh_token"];
                    credentials.ExpirationDate = tokenData.ContainsKey("expiration_date")
                                    ? tokenData["expiration_date"]
                                    : 0.0;
                }
                else
                {
                    throw new InvalidOperationException($"Error reading token file: {credentialsTokenPath}");
                }
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Error parsing JSON from token file: {credentialsTokenPath}", ex);
            }
            catch (KeyNotFoundException ex)
            {
                throw new InvalidOperationException($"Required key not found in token file: {credentialsTokenPath}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unexpected error reading token file: {credentialsTokenPath}", ex);
            }
        }
        private void ReadCredentialsApiKey(string credentialsApiKeyPath)
        {
            try
            {
                if (File.Exists(credentialsApiKeyPath))
                {
                    string apiKeyContent = File.ReadAllText(credentialsApiKeyPath);
                    var apiKeyData = JsonConvert.DeserializeObject<Dictionary<string, string>>(apiKeyContent);

                    if (apiKeyData == null)
                    {
                        throw new InvalidOperationException("Error reading apiKey file");
                    }

                    credentials.ApiKey = apiKeyData.ContainsKey("your_api_key") ? apiKeyData["your_api_key"] : "api_key_not_found";
                }
                else
                {
                    throw new FileNotFoundException($"File not found {credentialsApiKeyPath}");
                }
            }
            catch (FileNotFoundException ex)
            {
                throw new InvalidOperationException($"Error reading apiKey file: {credentialsApiKeyPath}. File not found.", ex);
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Error parsing JSON from apiKey file: {credentialsApiKeyPath}", ex);
            }
            catch (KeyNotFoundException ex)
            {
                throw new InvalidOperationException($"Required key not found in apiKey file: {credentialsApiKeyPath}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unexpected error reading apiKey file: {credentialsApiKeyPath}", ex);
            }
        }
        private string UpdateTokenAccess()
        {
            var body = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("client_id", credentials.ClientId),
                new KeyValuePair<string, string>("client_secret", credentials.ClientSecret),
                new KeyValuePair<string, string>("refresh_token", credentials.RefreshToken),
                new KeyValuePair<string, string>("grant_type", "refresh_token")
            });

            var headers = new Dictionary<string, string>
            {
                { "Content-Type", "application/x-www-form-urlencoded" }
            };

            string content = SendRequest(HttpMethod.Post, TOKEN_END_POINT, body, headers);

            return content;
        }
        private bool IsValidToken()
        {
            DateTime expirationDate = DateTime.FromOADate(credentials.ExpirationDate);
            DateTime now = DateTime.Now;

            return now >= expirationDate
                    ? false
                    : true;

        }
        private bool SaveToken(string content, string credentialsTokenPath)
        {
            try
            {
                var tokenData = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(content);

                if (tokenData == null)
                {
                    throw new InvalidOperationException("Error reading token data");
                }

                // Obtener el valor de expires_in y restar 5 minutos
                long expiresIn = tokenData["expires_in"];
                expiresIn -= 300;

                TimeSpan expirationTimeSpan = TimeSpan.FromSeconds(expiresIn);

                // Calcular la fecha de expiración
                DateTime expirationDate = DateTime.Now.Add(expirationTimeSpan);
                double expirationDateOADate = expirationDate.ToOADate();

                // Guardar la fecha de expiración en el diccionario
                tokenData["expiration_date"] = expirationDateOADate;

                if (!tokenData.ContainsKey("refresh_token"))
                {
                    tokenData["refresh_token"] = credentials.RefreshToken;
                }

                content = JsonConvert.SerializeObject(tokenData, Formatting.Indented);
                File.WriteAllText(credentialsTokenPath, content);

                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error saving token: " + ex.Message);
            }
        }
        private string ExchangeCodeForToken(string code)
        {
            var body = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("code", code),
                new KeyValuePair<string, string>("client_id", credentials.ClientId),
                new KeyValuePair<string, string>("client_secret", credentials.ClientSecret),
                new KeyValuePair<string, string>("redirect_uri", credentials.RedirectUri),
                new KeyValuePair<string, string>("grant_type", "authorization_code")
            });

            var headers = new Dictionary<string, string>
            {
                { "Content-type", "application/x-www-form-urlencoded" }
            };

            return SendRequest(HttpMethod.Post, TOKEN_END_POINT, body, headers);
        }
        private string WaitForCode()
        {
            string code = "";
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add(credentials.RedirectUri);
            listener.Start();

            try
            {
                HttpListenerContext context = listener.GetContext();
                HttpListenerRequest request = context.Request;
                code = request.QueryString["code"] ?? string.Empty;

                HttpListenerResponse response = context.Response;
                string responseString;

                if (!string.IsNullOrEmpty(code))
                {
                    responseString = "<!DOCTYPE html>" +
                                     "<html lang=\"es\">" +
                                     "<head>" +
                                     "<meta charset=\"UTF-8\" />" +
                                     "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />" +
                                     "<title>Document</title>" +
                                     "<style>" +
                                     "body { margin: 0; padding: 0; height: 100vh; background-color: #2d2d2d; backdrop-filter: blur(10px); opacity: 0.9; }" +
                                     ".code-editor { background-color: rgba(255, 255, 255, 0.8); color: #000000; font-family: 'Courier New', Courier, monospace; padding: 20px; border-radius: 5px; margin: 20px; backdrop-filter: blur(5px); }" +
                                     "</style>" +
                                     "</head>" +
                                     "<body>" +
                                     "<div class=\"code-editor\">" +
                                     "Código de autorización recibido. Puedes cerrar esta ventana." +
                                     "</div>" +
                                     "</body>" +
                                     "</html>";
                }
                else
                {
                    responseString = "<!DOCTYPE html>" +
                                     "<html lang=\"es\">" +
                                     "<head>" +
                                     "<meta charset=\"UTF-8\" />" +
                                     "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />" +
                                     "<title>Document</title>" +
                                     "<style>" +
                                     "body { margin: 0; padding: 0; height: 100vh; background-color: #2d2d2d; backdrop-filter: blur(10px); opacity: 0.9; }" +
                                     ".code-editor { background-color: rgba(255, 255, 255, 0.8); color: #000000; font-family: 'Courier New', Courier, monospace; padding: 20px; border-radius: 5px; margin: 20px; backdrop-filter: blur(5px); }" +
                                     "</style>" +
                                     "</head>" +
                                     "<body>" +
                                     "<div class=\"code-editor\">" +
                                     "Acceso denegado. Puedes cerrar esta ventana." +
                                     "</div>" +
                                     "</body>" +
                                     "</html>";
                }

                byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
                response.ContentLength64 = buffer.Length;
                Stream output = response.OutputStream;
                output.Write(buffer, 0, buffer.Length);
                output.Close();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error waiting for authorization code: " + ex.Message);
            }
            finally
            {
                listener.Stop();
            }

            return code;
        }
        private string GenerateOauthURL(string scopes)
        {
            var queryParams = new Dictionary<string, string>
                {
                    { "client_id", credentials.ClientId },
                    { "redirect_uri", credentials.RedirectUri },
                    { "response_type", "code" },
                    { "scope", scopes },
                    { "access_type", "offline" },
                    { "include_granted_scopes", "true" },
                    { "prompt", "consent" }
                };

            var queryString = string.Join("&", queryParams.Select(kvp => $"{kvp.Key}={Uri.EscapeDataString(kvp.Value)}"));
            return $"{OAUTH2_END_POINT}?{queryString}";
        }
        private string RequestToken(string scopes)
        {
            try
            {
                string url = GenerateOauthURL(scopes);
                System.Diagnostics.Process.Start(Navigator, url);

                string code = WaitForCode();

                if (string.IsNullOrEmpty(code)) throw new InvalidOperationException("Authorization code not received.");

                string jsonResponse = ExchangeCodeForToken(code);

                return Status == 200
                    ? jsonResponse
                    : "";
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error requesting token: " + ex.Message);
            }

        }
        private string SendRequest(HttpMethod method, string url, HttpContent body, Dictionary<string, string> headers)
        {
            using (var client = new HttpClient())
            {
                var request = new HttpRequestMessage(method, url) { Content = body };

                foreach (var header in headers)
                {
                    if (header.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
                    {
                        request.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(header.Value);
                    }
                    else
                    {
                        request.Headers.Add(header.Key, header.Value);
                    }
                }

                var response = client.SendAsync(request).Result;
                Status = (int)response.StatusCode;

                return response.Content.ReadAsStringAsync().Result;
            }
        }
        private class OAuthCredentials
        {
            public string ApiKey { get; set; } = string.Empty;
            public string AccessToken { get; set; } = string.Empty;
            public string ClientId { get; set; } = string.Empty;
            public string RedirectUri { get; set; } = string.Empty;
            public string ClientSecret { get; set; } = string.Empty;
            public string RefreshToken { get; set; } = string.Empty;
            public double ExpirationDate { get; set; } = 0.0;
        }
    }
}
