using GoogleApis.Drive.Interfaces;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using GoogleApis.Drive.Common;  

namespace GoogleApis.Drive.Component
{
    //// Implementar la interfaz IGoogleDrive en la clase GoogleDrive
    [ComVisible(true)]
    [Guid("f3fbc540-6c6d-4f3f-8f25-83337d0a73d6")] // Un GUID único para la clase
    [ClassInterface(ClassInterfaceType.None)] // Evita que se genere automáticamente una interfaz COM
    public class Files :GoogleDriveBase, IFiles
    {
        public Files(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }
        public int Operation()
        {
            return _status;
        }        
        private string GetNameFileForId(string fileID)
        {
            var queryParameters = new Scripting.Dictionary();
            queryParameters.Add("fields", "name");

            string json = GetMetadata(fileID, queryParameters);

            if (_status != 200)
            {
                throw new Exception("No response data");
            }

            var fileObject = JObject.Parse(json);
            string nameFile = fileObject.ContainsKey("name")
                ? fileObject["name"].ToString()
                : "unknown";

            return nameFile;
        }
        public string Copy(string fileID, string fileObject, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}/copy";
                string url = null;

                var headers = new Scripting.Dictionary();
                headers.Add("Content-Type", "application/json");

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, fileObject, headers);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Copy {fileID} : {ex.Message}");
            }
        }
        public string Delete(string fileID)
        {
            try
            {
                string pathParameters = $"/{fileID}";
                string url = null;

                url = CreateQueryParameters(pathParameters);

                var response = Request("DELETE", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Delete {fileID} : {ex.Message}");
            }
        }
        public string EmptyTrash(Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = "/trash";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("DELETE", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error EmptyTrash : {ex.Message}");
            }
        }
        public bool Export(string fileID, Scripting.Dictionary queryParameters, string pathFile)
        {
            try
            {
                string pathParameters = $"/{fileID}/export";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var buffer = Request("GET", url).Content.ReadAsByteArrayAsync();

                buffer.Wait();

                if (buffer.Result != null)
                {
                    System.IO.File.WriteAllBytes(pathFile, buffer.Result);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Export {fileID} : {ex.Message}");
            }
        }
        public string GenerateIds(Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = "/generateIds";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Generate Ids : {ex.Message}");
            }
        }
        public string GetMetadata(string fileID, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error get metadata of {fileID} : {ex.Message}");
            }
        }
        public string List(Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string url = null;

                url = CreateQueryParameters(queryParameters: queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error list : {ex.Message}");
            }
        }
        public string ListLabels(string fileID, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}/listLabels";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error list labels {fileID} : {ex.Message}");
            }
        }
        public string Update(string fileID, string fileObject = null, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("PATCH", url, fileObject);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error update {fileID} : {ex.Message}");
            }
        }
        public bool DownLoadContentLink(string fileID)
        {
            const string WEB_CONTENT_LINK = "webContentLink";
            var fields = new Scripting.Dictionary();
            string webContentLink = null;
            string json = null;
            try
            {
                //fields["fields"] = WEB_CONTENT_LINK;
                fields.Add("fields", WEB_CONTENT_LINK);
                json = GetMetadata(fileID, fields);

                if (json != null)
                {
                    var fileObject = JObject.Parse(json);

                    if (fileObject[WEB_CONTENT_LINK] != null)
                    {
                        webContentLink = fileObject[WEB_CONTENT_LINK].ToString();
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = webContentLink,
                            UseShellExecute = true,
                        });
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error download for content link {fileID} : {ex.Message}");
            }
        }
        public bool DownLoad(string fileID, string directory)
        {
            if (!System.IO.Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException($"Not found directory {directory}");
            }

            string nameFile = GetNameFileForId(fileID);
            string fullPath = Path.Combine(directory, nameFile);
            var queryParameters = new Scripting.Dictionary();

            queryParameters.Add("alt", "media");
            string pathParameters = $"/{fileID}";

            string url = CreateQueryParameters(pathParameters, queryParameters);
            var buffer = Request("GET", url).Content.ReadAsByteArrayAsync();

            buffer.Wait();

            if (buffer.Result != null)
            {
                System.IO.File.WriteAllBytes(fullPath, buffer.Result);
                return true;
            }
            return false;
        }
        public string UploadMedia(string pathFile = null, string mimeType = "application/octet-stream")
        {
            bool endPointUpload = false;
            string body = null;
            byte[] buffer = null;
            var headers = new Scripting.Dictionary();
            HttpResponseMessage response = null;

            if (pathFile != null)
            {
                if (System.IO.File.Exists(pathFile))
                {
                    endPointUpload = true;
                    buffer = Helper.Helper.SourceToBinary(pathFile);
                    var fileInfo = new FileInfo(pathFile);

                    headers.Add("Content-Type", mimeType);
                    headers.Add("Content-Length", fileInfo.Length.ToString());
                    headers.Add("Content-Transfer-Encoding", "binary");
                }
                else
                {
                    throw new FileNotFoundException($"File not exists {pathFile}");
                }
            }
            else
            {
                var fileObject = new
                {
                    mimeType = "application/vnd.google-apps.folder",
                    parents = "root"
                };

                body = JsonConvert.SerializeObject(fileObject);
            }

            var queryParameters = new Scripting.Dictionary();

            queryParameters.Add("uploadType", "media");

            string url = CreateQueryParameters(queryParameters: queryParameters,
                                                endPointUpload: endPointUpload);

            response = body != null
                ? Request("POST", url, body, headers)
                : Request("POST", url, buffer, headers);

            return response.Content.ReadAsStringAsync().Result;
        }
        public string UploadMultipart(string pathFile, string fileObject)
        {
            try
            {
                string boundary = Helper.Helper.GenerateString(15);
                var queryParameters = new Scripting.Dictionary {
                    {"uploadType", "multipart" }
                };

                string url = CreateQueryParameters(queryParameters: queryParameters, endPointUpload: true);

                using (var client = new HttpClient())
                {
                    var requestMessage = new HttpRequestMessage(HttpMethod.Post, url);
                    string body = Helper.Helper.CreatePartRelated(pathFile, boundary, fileObject);

                    requestMessage.Headers.TryAddWithoutValidation("Authorization", $"Bearer {_accessToken}");
                    requestMessage.Headers.TryAddWithoutValidation("Accept", "application/json");
                    requestMessage.Content = new StringContent(body, Encoding.UTF8);
                    requestMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("multipart/related");
                    requestMessage.Content.Headers.ContentType.Parameters.Add(new NameValueHeaderValue("boundary", boundary));

                    var response = client.SendAsync(requestMessage).Result;
                    return response.Content.ReadAsStringAsync().Result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error upload {pathFile} : {ex.Message}");
            }
        }
    }
}
