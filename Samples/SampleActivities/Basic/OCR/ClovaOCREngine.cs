using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using UiPath.OCR.Contracts;
using UiPath.OCR.Contracts.Activities;
using UiPath.OCR.Contracts.DataContracts;

namespace SampleActivities.Basic.OCR
{
    [DisplayName("Clova OCR Engine")]
    public class ClovaOCREngine : OCRCodeActivity
    {
        [Category("Input")]
        [Browsable(true)]
        public override InArgument<Image> Image { get => base.Image; set => base.Image = value; }

        [Category("Login")]
        [RequiredArgument]
        [Description("Clova OCR API GW endpoint 情報")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Login")]
        [RequiredArgument]
        [Description("Clova OCR Secret")]
        public InArgument<string> Secret { get; set; }

        [Category("Option")]
        [Browsable(true)]
        [Description("利用可能な言語は ja ko zh-TW のいずれかです")]
        public InArgument<string> Languages { get; set; } = "ja";

        [Category("Option")]
        [Browsable(true)]
        [Description("表抽出オプション")]
        public InArgument<Boolean> enableTableDetection { get; set; } = false;


        [Category("Output")]
        [Browsable(true)]
        public override OutArgument<string> Text { get => base.Text; set => base.Text = value; }

        [Category("InOutput")]
        [Browsable(true)]
        public InOutArgument<DataSet> DataSet { get; set; }

        private string file_path;
        private DataSet dataSet;

        /**
         * OCRENgine が動作するために必要な関数の実装 
         * Dictionary<string,object> options に必要な値を格納して引き渡します。 
         */
        public override Task<OCRResult> PerformOCRAsync(Image image, Dictionary<string, object> options, CancellationToken ct)
        {

            file_path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "clova_ocr_req_image.png");
            if( image != null ) {
                if (System.IO.File.Exists(file_path))
                    System.IO.File.Delete(file_path);
#if DEBUG
                System.Console.WriteLine($"width={image.Width}, height={image.Height} resolution={image.HorizontalResolution} ");
#endif
                image.Save(file_path, System.Drawing.Imaging.ImageFormat.Png);
            } else
            {
                file_path = string.Empty;
            }
 #if DEBUG
            System.Console.WriteLine("temp file path " + file_path);
#endif

            var result =   ClovaOCRResultHelper.FromClovaClient(file_path, options);

            dataSet = (DataSet)options["dataset"];
            return result;
        }

        /**
         * Output 出力を設定します。PeformOCRAsync から options に含まれる値を利用して、最終的に Output argument へ値を設定します。 
         */
        protected override void OnSuccess(CodeActivityContext context, OCRResult result)
        {
            DataSet.Set(context, dataSet);
        }
        //protected override void on

        protected override Dictionary<string, object> BeforeExecute(CodeActivityContext context)
        {
            return new Dictionary<string, object>
            {
                { "endpoint",  Endpoint.Get(context) },
                { "secret", Secret.Get(context) },
                { "lang", Languages.Get(context) },
                { "enableTableDetection", enableTableDetection.Get(context) },
                { "dataset",  DataSet.Get(context) }
            };
        }
    }
}
