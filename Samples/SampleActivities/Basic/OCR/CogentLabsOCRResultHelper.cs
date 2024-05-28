using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using UiPath.OCR.Contracts.DataContracts;
using UiPath.OCR.Contracts;
using UiPath.DocumentProcessing.Contracts.Dom;
using Azure.Core;
using System.Threading;
using System.Diagnostics;

namespace SampleActivities.Basic.OCR
{
    internal static class CogentLabsOCRResultHelper
    {
        internal class RequestBody
        {
            public ExportSettings exportSettings { get; set; } = new ExportSettings { type = "json", aggregation = "oneFile" };
            public string name { get; set; } = "UiPath全文OCR抽出タスク";
            public bool allowUseOfData { get; set; } = false;
            public string description { get; set; } = "UiPathのカスタムアクティビティから実行しているタスクです。";
//            public Dictionary<string, string> labels { get; set; } = new Dictionary<string, string> { { "property1", "string" }, { "property2", "string" } };
            public string[] languages { get; set; } = new string[] { "ja" };
            public string requestType { get; set; } = "freeform";
        }

        internal class ExportSettings
        {
            public string type { get; set; }
            public string aggregation { get; set; }
        }

        internal static async Task<OCRResult> FromCogentLabsClient(string file_path, Dictionary<string, object> options)
        {
            OCRResult ocrResult = new OCRResult();
            var apiUrl = "https://api.smartread.jp/v3";
            var apiKey = options["apiKey"].ToString();

            var client = new CogentLabHttpClient(apiUrl);
            client.SetApiKey(apiKey);

            // Step 1: Create task
            var taskResponse = await CreateTask(client);
            var taskId = JsonConvert.DeserializeObject<dynamic>(await taskResponse.Content.ReadAsStringAsync()).taskId;
#if DEBUG
            Console.WriteLine($"taskId : {taskId}");
#endif

            // Step 2: Upload file
            String requestId;
            try
            {
                client.Clear();
                client.AddFile(file_path);
                var uploadResponse = await client.PostAsync($"/task/{taskId}/request");
                requestId = JObject.Parse(await uploadResponse.Content.ReadAsStringAsync())["requestId"].ToString();
#if DEBUG
                Console.WriteLine($"requestId: {requestId}");
                Console.WriteLine($"Headers: {uploadResponse.Headers}");
                Console.WriteLine($"ReasonPhrase: {uploadResponse.ReasonPhrase}");
                Console.WriteLine($"RequestMessage: {uploadResponse.RequestMessage}");
                Console.WriteLine($"StatusCode: {uploadResponse.StatusCode}");
#endif
                if (!uploadResponse.IsSuccessStatusCode)
                {
                    Console.WriteLine($"File upload failed: {uploadResponse.StatusCode}");
                    // 作成したタスクを削除
                    await DeleteTask(client, taskId.ToString());
                    throw new ApplicationException("ファイルアップロード時に何かしらのエラーが発生しました。");
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("実行時に何かしらのエラーが発生しました。", ex);
            }

            // Step 3: Get request ID and check status
            var status = await CheckRequestStatus(client, taskId.ToString());

            // Step 4: Get results
            var resultResponse = await GetRequestResults(client, requestId);
#if DEBUG
            Console.WriteLine($"resultResponse : {resultResponse}");
#endif
            var resultContent = await resultResponse.Content.ReadAsStringAsync();
#if DEBUG
            Console.WriteLine($"resultContent : {resultContent}");
#endif
            var ocrResults = JsonConvert.DeserializeObject<JObject>(resultContent);
#if DEBUG
            Console.WriteLine($"ocrResults : {ocrResults}");
#endif
            // Convert OCR results to OCRResult type
            ocrResult = ConvertToOCRResult(ocrResults);
#if DEBUG
            Console.WriteLine($"ocrResult : {ocrResult}");
#endif

            // Step 5: Delete task
            await DeleteTask(client, taskId.ToString());

            return ocrResult;
        }

        private static async Task<HttpResponseMessage> CreateTask(CogentLabHttpClient client)
        {
            var requestBody = new RequestBody();
            client.Clear();
            client.AddField("data", JsonConvert.SerializeObject(requestBody));
            return await client.PostAsync("/task", new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json"));
        }

        private static async Task<string> CheckRequestStatus(CogentLabHttpClient client, string taskId)
        {
            HttpResponseMessage statusResponse;
            string status;
            int loop_counter = 0;
            do
            {
                loop_counter++;
                statusResponse = await GetTaskStatus(client, taskId);
                var statusContent = await statusResponse.Content.ReadAsStringAsync();
                var statusResult = JsonConvert.DeserializeObject<dynamic>(statusContent);
                status = statusResult.state.ToString();
#if DEBUG
                Console.WriteLine($"Current status: {status}");
#endif
                if (statusResult.requestStateSummary.OCR_COMPLETED > 0)
                {
                    break;
                }
                else if ( loop_counter > 120 )  // 10分(5秒x120回=600秒)経過したらタイムアウトを発生
                {
                    throw new ApplicationException("OCR読取りに時間が掛かったため、タイムアウトが発生しました。");
                }
                await Task.Delay(5000);         // 5秒間スリープ
            } while (statusResponse.IsSuccessStatusCode && status != "COMPLETED");
            return status;
        }

        private static async Task<HttpResponseMessage> GetTaskStatus(CogentLabHttpClient client, string taskId)
        {
            return await client.GetAsync($"/task/{taskId}");
        }

        private static async Task<HttpResponseMessage> GetRequestResults(CogentLabHttpClient client, string requestId)
        {
            return await client.GetAsync($"/request/{requestId}/results?offset=0&limit=100");   // 最大100ページの制限を実施
        }

        private static async Task<HttpResponseMessage> DeleteTask(CogentLabHttpClient client, string taskId)
        {
            return await client.DeleteAsync($"/task/{taskId}");
        }

        private static OCRResult ConvertToOCRResult(JObject ocrResults)
        {
            var words = new List<UiPath.OCR.Contracts.DataContracts.Word>();

            foreach (var result in ocrResults["results"])
            {
                foreach (var page in result["pages"])
                {
                    foreach (var field in page["fields"])
                    {
                        var singleLine = field["singleLine"];
                        var boundingBox = field["boundingBox"];

                        var word = new UiPath.OCR.Contracts.DataContracts.Word
                        {
                            Text = singleLine["text"].ToString(),
                            Confidence = (int)(singleLine["confidence"].ToObject<double>() * 100),
                            PolygonPoints = new PointF[]
                            {
                        new PointF((float)boundingBox["x"], (float)boundingBox["y"]),
                        new PointF((float)boundingBox["x"] + (float)boundingBox["width"], (float)boundingBox["y"]),
                        new PointF((float)boundingBox["x"] + (float)boundingBox["width"], (float)boundingBox["y"] + (float)boundingBox["height"]),
                        new PointF((float)boundingBox["x"], (float)boundingBox["y"] + (float)boundingBox["height"])
                            },
                            Characters = singleLine["characters"].Select(ch => new Character
                            {
                                Char = ch["character"].ToString()[0],
                                Confidence = (int)(ch["confidence"].ToObject<double>() * 100),
                                PolygonPoints = new PointF[]
                                {
                            new PointF((float)ch["boundingBox"]["x"], (float)ch["boundingBox"]["y"]),
                            new PointF((float)ch["boundingBox"]["x"] + (float)ch["boundingBox"]["width"], (float)ch["boundingBox"]["y"]),
                            new PointF((float)ch["boundingBox"]["x"] + (float)ch["boundingBox"]["width"], (float)ch["boundingBox"]["y"] + (float)ch["boundingBox"]["height"]),
                            new PointF((float)ch["boundingBox"]["x"], (float)ch["boundingBox"]["y"] + (float)ch["boundingBox"]["height"])
                                }
                            }).ToArray()
                        };
                        words.Add(word);
                    }
                }
            }

            return new OCRResult
            {
                Text = string.Join(" ", words.Select(w => w.Text)),
                Words = words.ToArray()
            };
        }
    }
}
