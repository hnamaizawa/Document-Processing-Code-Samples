using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace SampleActivities.Basic.OCR
{
    public class AIinsideHttpClient
    {
        public AIinsideHttpClient(string endpoint)
        {
            this.url = endpoint;
            this.client = new HttpClient();
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public void SetApiKey(string apiKey)
        {
            this.client.DefaultRequestHeaders.Add("apikey", apiKey);
        }

        public void AddFile(string fileName, string fieldName = "file")
        {
            var fstream = System.IO.File.OpenRead(fileName);
#if DEBUG
            Console.WriteLine($"AddFile() file size: {fstream.Length}");
            Console.WriteLine($"AddFile() fieldName : {fieldName}");
            Console.WriteLine($"AddFile() fileName : {fileName}");
#endif
            var fileContent = new StreamContent(fstream);
            fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
            {
                Name = fieldName,
                FileName = Path.GetFileName(fileName)
            };

            var extension = Path.GetExtension(fileName).ToLower();
            if (extension == ".pdf")
            {
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
            }
            else if (extension == ".jpg" || extension == ".jpeg")
            {
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/jpeg");
            }
            else if (extension == ".png")
            {
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");
            }
            else if (extension == ".tiff" || extension == ".tif")
            {
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/tiff");
            }
            else
            {
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");
            }

            this.content.Add(fileContent);
        }

        public void AddField(string name, string value)
        {
            this.content.Add(new StringContent(value), name);
        }

        public string GetUrlWithParam(string relativeUrl, string name, string value)
        {
            if (!string.IsNullOrEmpty(this.url) && !this.url.Contains("?"))
            {
                return relativeUrl + $"?{name}={value}";
            }
            else
            {
                return relativeUrl + $"&{name}={value}";
            }
        }

        public void Clear()
        {
            this.content.Dispose();
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public async Task<HttpResponseMessage> PostAsync(string relativeUrl)
        {
#if DEBUG
            Console.WriteLine($"PostAsync() URL : {this.url + relativeUrl}");
#endif
            return await this.client.PostAsync(this.url + relativeUrl, this.content);
        }

        public async Task<HttpResponseMessage> PostAsync(string relativeUrl, HttpContent content)
        {
#if DEBUG
            Console.WriteLine($"PostAsync() URL : {this.url + relativeUrl}");
            Console.WriteLine($"PostAsync() Content : {content.ReadAsStringAsync().Result}");
#endif
            return await this.client.PostAsync(this.url + relativeUrl, content);
        }

        public async Task<HttpResponseMessage> GetAsync(string relativeUrl)
        {
#if DEBUG
            Console.WriteLine($"GetAsync() URL : {this.url + relativeUrl}");
#endif
            return await this.client.GetAsync(this.url + relativeUrl);
        }

        public async Task<HttpResponseMessage> DeleteAsync(string relativeUrl)
        {
#if DEBUG
            Console.WriteLine($"DeleteAsync() URL : {this.url + relativeUrl}");
#endif
            return await this.client.DeleteAsync(this.url + relativeUrl);
        }

        public MultipartFormDataContent GetContent()
        {
            return this.content;
        }

        private HttpClient client;
        private string url;
        private MultipartFormDataContent content;
    }
}
