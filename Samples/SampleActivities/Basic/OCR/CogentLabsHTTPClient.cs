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
    public class CogentLabHttpClient
    {
        public CogentLabHttpClient(string endpoint)
        {
            this.url = endpoint;
            this.client = new HttpClient();
            this.content = new MultipartFormDataContent("ocr----" + DateTime.Now.Ticks.ToString());
        }

        public void SetApiKey(string apiKey)
        {
            this.client.DefaultRequestHeaders.Add("Authorization", "apikey " + apiKey);
        }

        public void AddFile(string fileName, string fieldName = "image")
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
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/pdf");
            this.content.Add(fileContent);
        }

        public void AddField(string name, string value)
        {
            this.content.Add(new StringContent(value), name);
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

        private HttpClient client;
        private string url;
        private MultipartFormDataContent content;
    }
}
