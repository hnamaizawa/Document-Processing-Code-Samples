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
    internal static class AIinsideOCRResultHelper
    {
        internal static async Task<OCRResult> FromAIinsideClient(string file_path, Dictionary<string, object> options)
        {
            var apiKey = options["apiKey"].ToString();
            var client = new AIinsideHttpClient(options["endpoint"].ToString());
            string registerId;
            HttpResponseMessage resultResponse;
            JObject ocrResults;
            OCRResult ocrResult;
            client.SetApiKey(apiKey);

            // Step 1: �o�^ API /register
            try
            {
                client.Clear();
                client.AddFile(file_path);
                client.AddField("concatenate", "0");
                client.AddField("characterExtraction", "1");
                client.AddField("tableExtraction", "1");
                var registerResponse = await client.PostAsync("/register", client.GetContent());
                if (!registerResponse.IsSuccessStatusCode)
                {
                    throw new ApplicationException("�o�^���N�G�X�g�����s���܂����B");
                }
                registerId = JsonConvert.DeserializeObject<dynamic>(await registerResponse.Content.ReadAsStringAsync()).id;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("�o�^���N�G�X�g���ɃG���[���������܂����B", ex);
            }

            // Step 2: ���ʎ擾 API /getOcrResult
            try
            {
                resultResponse = await CheckRequestStatus(client, registerId.ToString());
#if DEBUG
                Console.WriteLine($"resultResponse : {resultResponse}");
#endif
            }
            catch (Exception ex)
            {
                throw new ApplicationException("���ʎ擾���N�G�X�g���ɃG���[���������܂����B", ex);
            }

            // Step 3: Get results
            try
            {
                var resultContent = await resultResponse.Content.ReadAsStringAsync();
#if DEBUG
                Console.WriteLine($"resultContent : {resultContent}");
#endif
                ocrResults = JsonConvert.DeserializeObject<JObject>(resultContent);
#if DEBUG
                Console.WriteLine($"ocrResults : {ocrResults}");
#endif
            }
            catch (Exception ex)
            {
                throw new ApplicationException("���ʂ̉�͒��ɃG���[���������܂����B", ex);
            }

            // Step 4: Convert OCR results to OCRResult type
            try
            {
                var img = System.Drawing.Image.FromFile(file_path);
                var width = img.Width;
                var height = img.Height;

                ocrResult = ConvertToOCRResult(ocrResults, width, height);
#if DEBUG
                Console.WriteLine($"ocrResult : {ocrResult}");
#endif
            }
            catch (Exception ex)
            {
                throw new ApplicationException("OCR���ʂ̕ϊ����ɃG���[���������܂����B", ex);
            }

            // Step 5: �폜 API /delete
            try
            {
                var deleteRequest = new { fullOcrJobId = registerId };
                var content = new StringContent(JsonConvert.SerializeObject(deleteRequest), Encoding.UTF8, "application/json");
                var response = await client.PostAsync("/delete", content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new ApplicationException("�폜���N�G�X�g�����s���܂����B");
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("�폜���N�G�X�g���ɃG���[���������܂����B", ex);
            }

            Console.WriteLine($"DBG ocrResult: Text={ocrResult.Text}, Length ={ocrResult.Words.Length}");

            return ocrResult;
        }

        private static async Task<HttpResponseMessage> CheckRequestStatus(AIinsideHttpClient client, string taskId)
        {
            HttpResponseMessage statusResponse;
            string status;
            int loop_counter = 0;
            do
            {
                loop_counter++;
                statusResponse = await GetTaskStatus(client, taskId);
                if (!statusResponse.IsSuccessStatusCode)
                {
                    throw new ApplicationException("�X�e�[�^�X�擾���N�G�X�g�����s���܂����B");
                }
                status = JsonConvert.DeserializeObject<dynamic>(await statusResponse.Content.ReadAsStringAsync()).status;
#if DEBUG
                Console.WriteLine($"Current status: {status}");
#endif
                if (status == "done")
                {
                    break;
                }
                else if (loop_counter > 120)  // 10�� (5�b x 120�� = 600�b) �o�߂�����^�C���A�E�g
                {
                    throw new ApplicationException("OCR�������������Ȃ����߁A�^�C���A�E�g���܂����B");
                }
                await Task.Delay(5000);  // 5�b�ԃX���[�v
            } while (status != "done");
            return statusResponse;
        }

        private static async Task<HttpResponseMessage> GetTaskStatus(AIinsideHttpClient client, string taskId)
        {
            client.Clear();
            var urlWithParam = client.GetUrlWithParam("/getOcrResult", "id", taskId);
            return await client.GetAsync(urlWithParam);
        }

        private static OCRResult ConvertToOCRResult(JObject ocrResults, int width, int height)
        {
            var words = new List<UiPath.OCR.Contracts.DataContracts.Word>();

            foreach (var result in ocrResults["results"])
            {
                foreach (var page in result["pages"])
                {
                    foreach (var field in page["ocrResults"])
                    {
                        var singleLine = field["text"];
                        var boundingBox = field["bbox"];

                        var word = new UiPath.OCR.Contracts.DataContracts.Word
                        {
                            Text = singleLine.ToString(),
                            Confidence = 100,
                            PolygonPoints = new PointF[]
                            {
                                new PointF((float)boundingBox["left"] * width, (float)boundingBox["top"] * height),
                                new PointF((float)boundingBox["right"] * width, (float)boundingBox["top"] * height),
                                new PointF((float)boundingBox["right"] * width, (float)boundingBox["bottom"] * height),
                                new PointF((float)boundingBox["left"] * width, (float)boundingBox["bottom"] * height)
                            },
                            Characters = field["characters"].Select(ch => new Character
                            {
                                Char = ch["char"].ToString()[0],
                                Confidence = (int)(ch["ocrConfidence"].ToObject<double>() * 100),
                                PolygonPoints = new PointF[]
                                {
                                    new PointF((float)ch["bbox"]["left"] * width, (float)ch["bbox"]["top"] * height),
                                    new PointF((float)ch["bbox"]["right"] * width, (float)ch["bbox"]["top"] * height),
                                    new PointF((float)ch["bbox"]["right"] * width, (float)ch["bbox"]["bottom"] * height),
                                    new PointF((float)ch["bbox"]["left"] * width, (float)ch["bbox"]["bottom"] * height)
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
