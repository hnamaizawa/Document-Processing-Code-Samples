using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleActivities.Basic.OCR;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UiPath.OCR.Contracts.DataContracts;
using UiPath.OCR.Contracts.Activities;
using UiPath.OCR.Contracts;
using Microsoft.Extensions.Configuration;

namespace SampleActivities.Basic.OCR.Tests
{
    [TestClass()]
    public class AIinsideTests
    {
        [TestMethod()]
        public async Task AIinsideTestsSuccess()
        {
            // arange
            var filePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\File\\"));
            var settings = ConfigurationHelper.GetOCRSettings();
            var imagePath = Path.Combine(filePath, "Sample.png");
            var image = Image.FromFile(imagePath);

            var options = new Dictionary<string, object>
            {
                { "endpoint", settings.AIinside.Endpoint },
                { "apiKey", settings.AIinside.ApiKey }
            };
            var ocr = new AIinsideOCREngine();
            var ct = new CancellationToken();

            // act
            var result = await ocr.PerformOCRAsync(image, options, ct);

            // assert
            Assert.IsTrue(result.Text.Contains("東京都"));
        }
    }

    public static class ConfigurationHelper
    {
        public static OCRSettings GetOCRSettings()
        {
            var basePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\"));
            var configuration = new ConfigurationBuilder()
                .SetBasePath(basePath)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            return configuration.GetSection("OCRSettings").Get<OCRSettings>();
        }
    }

    public class OCRSettings
    {
        public AIinsideSettings AIinside { get; set; }= new AIinsideSettings();
    }

    public class AIinsideSettings
    {
        public string? Endpoint { get; set; }
        public string? ApiKey { get; set; }

    }

}