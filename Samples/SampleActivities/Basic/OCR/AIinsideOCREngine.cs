using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.OCR.Contracts;
using UiPath.OCR.Contracts.Activities;
using UiPath.OCR.Contracts.DataContracts;
using static System.Net.Mime.MediaTypeNames;

namespace SampleActivities.Basic.OCR
{
    [DisplayName("AI inside OCR Engine")]
    public class AIinsideOCREngine : OCRCodeActivity
    {
        [Category("Input")]
        [Browsable(true)]
        public override InArgument<System.Drawing.Image> Image { get => base.Image; set => base.Image = value; }

        [Category("Login")]
        [RequiredArgument]
        [Description("AI inside API GW endpoint")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]
        [RequiredArgument]
        [Description("AI inside OCR ApiKey")]
        public InArgument<string> ApiKey { get; set; }

        [Category("Output")]
        [Browsable(true)]
        public override OutArgument<string> Text { get => base.Text; set => base.Text = value; }

        private string file_path;

        /**
         * OCR Engine が動作するために必要な関数の実装 
         * Dictionary<string,object> options に必要な値を格納して引き渡します。 
         */
        public override Task<OCRResult> PerformOCRAsync(System.Drawing.Image image, Dictionary<string, object> options, CancellationToken ct)
        {
            file_path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "AI_inside_ocr_req_image." + DateTime.Now.Ticks.ToString() + ".png");
            if ( image != null ) {
                if (System.IO.File.Exists(file_path))
                    System.IO.File.Delete(file_path);
#if DEBUG
                Console.WriteLine($"width={image.Width}, height={image.Height} resolution={image.HorizontalResolution} ");
#endif
                image.Save(file_path, System.Drawing.Imaging.ImageFormat.Png);
            } else
            {
                file_path = string.Empty;
            }
#if DEBUG
            Console.WriteLine("temp file path " + file_path);
#endif
            var result = AIinsideOCRResultHelper.FromAIinsideClient(file_path, options);

            return result;
        }

        /**
         * Output 出力を設定します。PeformOCRAsync から options に含まれる値を利用して、最終的に Output argument へ値を設定します。 
         */
        protected override void OnSuccess(CodeActivityContext context, OCRResult result)
        {
            var Words = result.Words;
            var Text = Words.Select(word => word.Text);
            var Confidence = Words.Select(word => word.Confidence);
            var PolygonPoints = Words.Select(word => word.PolygonPoints);
            var Characters = Words.Select(word => word.Characters);
            Console.WriteLine($"Text={Text}, Confidence={Confidence} PolygonPoints={PolygonPoints} PolygonPoints={Characters}");


        }
        //protected override void on

        protected override Dictionary<string, object> BeforeExecute(CodeActivityContext context)
        {
            return new Dictionary<string, object>
            {
                { "endpoint",  Endpoint.Get(context) },
                { "apiKey", ApiKey.Get(context) },
            };
        }
    }
}
