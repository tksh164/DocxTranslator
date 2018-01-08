using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace DocxTranslator
{
    internal class TranslatorClient : IDisposable
    {
        public string ApiKey { get; private set; }
        public string AuthorizationToken { get; private set; }

        private static HttpClient TranslatorApiHttpClient;  // for api.microsofttranslator.com

        private TranslatorClient()
        {
            TranslatorApiHttpClient = new HttpClient();
        }

        public static TranslatorClient Create(string apiKey)
        {
            var client = new TranslatorClient()
            {
                ApiKey = apiKey,
            };

            var authToken = GetAuthorizationTokenAsync(apiKey);
            authToken.Wait();
            client.AuthorizationToken = authToken.Result;

            return client;
        }

        public void Dispose()
        {
            TranslatorApiHttpClient.Dispose();
        }

        private static async Task<string> GetAuthorizationTokenAsync(string apiKey)
        {
            const string requestUri = "https://api.cognitive.microsoft.com/sts/v1.0/issueToken";

            try
            {
                using (var httpClient = new HttpClient())
                using (var request = new HttpRequestMessage())
                {
                    request.Method = HttpMethod.Post;
                    request.RequestUri = new Uri(requestUri);
                    request.Headers.Add("Accept", "application/jwt");
                    request.Headers.Add("Ocp-Apim-Subscription-Key", apiKey);

                    var response = await httpClient.SendAsync(request);
                    var responseBody = await response.Content.ReadAsStringAsync();

                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            return responseBody;

                        default:
                            throw new Exception("HTTP response error");
                    }
                }
            }
            catch (WebException ex)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static async Task<Dictionary<string, (string displayName, string direction)>> GetTextTranslatableLanguagesAsync()
        {
            const string requestUri = "https://dev.microsofttranslator.com/languages?api-version=1.0&scope=text";

            try
            {
                using (var httpClient = new HttpClient())
                using (var request = new HttpRequestMessage())
                {
                    request.Method = HttpMethod.Get;
                    request.RequestUri = new Uri(requestUri);
                    request.Headers.Add("Accept", "application/json");

                    var response = await httpClient.SendAsync(request);
                    var responseBody = await response.Content.ReadAsStringAsync();

                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:

                            var result = new Dictionary<string, (string displayName, string direction)>();

                            dynamic langObj = JObject.Parse(responseBody)["text"].First;
                            while (langObj != null)
                            {
                                string name = langObj.Name;
                                string displayName = langObj.Value["name"];
                                string direction = langObj.Value["dir"];

                                result.Add(name, (displayName: displayName, direction: direction));

                                langObj = langObj.Next;
                            }

                            return result;

                        default:
                            throw new Exception("HTTP response error");
                    }
                }
            }
            catch (WebException ex)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw;
            };
        }

        public async Task<string[]> TranslateTextArray(string translationFrom, string translationTo, string[] translationTexts)
        {
            const string requestUri = "https://api.microsofttranslator.com/V2/Http.svc/TranslateArray";

            try
            {
                using (var request = new HttpRequestMessage())
                {
                    //
                    // Create the request.
                    //

                    request.Method = HttpMethod.Post;
                    request.RequestUri = new Uri(requestUri);
                    request.Headers.Add("Authorization", string.Format("Bearer {0}", AuthorizationToken));

                    var requestBody = CreateTranslateArrayRequestBody(translationTo, translationTexts, translationFrom);
                    request.Content = new StringContent(requestBody, Encoding.UTF8, "application/xml");

                    //
                    // Send the request and get the response.
                    //

                    var response = await TranslatorApiHttpClient.SendAsync(request);
                    var responseBody = await response.Content.ReadAsStringAsync();

                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:

                            var result = new List<string>();
                            var deserializedObj = DeserializeTranslatedResponse(responseBody);
                            for (int i = 0; i < deserializedObj.TranslateArrayResponse.Length; i++)
                            {
                                result.Add(deserializedObj.TranslateArrayResponse[i].TranslatedText);
                            }
                            return result.ToArray();

                        default:
                            throw new Exception("HTTP response error");
                    }
                }
            }
            catch (WebException ex)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static string CreateTranslateArrayRequestBody(string translateTo, string[] translationTexts, string translateFrom = "", string appId = "", string category = "generalnn", string contentType = "text/plain", string state = "", string filterUri = "all", string filterUser = "all", string profanityAction = "NoAction")
        {
            var translateArrayRequest = new TranslateArrayRequestElement()
            {
                AppId = appId,
                From = translateFrom,
                Options = new OptionsElement()
                {
                    Category = category,
                    ContentType = contentType,
                    ReservedFlags = string.Empty,
                    State = state,
                    Uri = filterUri,
                    User = filterUser,
                    ProfanityAction = profanityAction,
                },
                Texts = translationTexts,
                To = translateTo,
            };

            var stringBuilder = new StringBuilder();
            var xmlWriterSettings = new XmlWriterSettings()
            {
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = true,
                Indent = true,
                IndentChars = "    ",
                NewLineChars = "\r\n",
                NewLineHandling = NewLineHandling.Replace,
            };
            using (var xmlWriter = XmlWriter.Create(stringBuilder, xmlWriterSettings))
            {
                var ns = new XmlSerializerNamespaces();
                ns.Add(string.Empty, string.Empty);
                var xmlSerializer = new XmlSerializer(typeof(TranslateArrayRequestElement));

                xmlSerializer.Serialize(xmlWriter, translateArrayRequest, ns);
            }

            return stringBuilder.ToString();
        }

        private static ArrayOfTranslateArrayResponseElement DeserializeTranslatedResponse(string responseBody)
        {
            var xmlReaderSettings = new XmlReaderSettings()
            {
                DtdProcessing = DtdProcessing.Ignore,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                ValidationType = ValidationType.None,
            };

            using (var reader = new StringReader(responseBody))
            using (var xmlReader = XmlReader.Create(reader, xmlReaderSettings))
            {
                var xmlSerializer = new XmlSerializer(typeof(ArrayOfTranslateArrayResponseElement));
                return (ArrayOfTranslateArrayResponseElement)xmlSerializer.Deserialize(xmlReader);
            }
        }
    }
}
