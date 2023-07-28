using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.Xml.Linq;
using Microsoft.AspNetCore.Html;
using System.Text;
using Aspose.Words;

namespace HtmlToDocx.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        [Route("/OpenXmlPowerTools")]
        public IActionResult ConvertHtmlToDocx([FromBody] string html)
        {
            HtmlToWmlConverterSettings settings = new HtmlToWmlConverterSettings();

            XElement htmlAsXElement = XElement.Parse(html);

            WmlDocument convertedDocument = HtmlToWmlConverter.ConvertHtmlToWml(
                "",
                "",
                "",
                htmlAsXElement,
                settings);

            byte[] byteArray = convertedDocument.DocumentByteArray;
            return File(byteArray, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
        }

        [HttpPost]
        [Route("/Aspose")]
        public async Task<IActionResult> ConvertHtmlToDocxAspose([FromForm] IFormFile htmlFile)
        {
            byte[] result;
            using (MemoryStream stream = new MemoryStream())
            {
                await htmlFile.CopyToAsync(stream);
                stream.Seek(0, SeekOrigin.Begin);
                Document doc = new Document(stream);
                using (MemoryStream outputStream = new MemoryStream())
                {
                    doc.Save(outputStream, SaveFormat.Docx);
                    outputStream.Seek(0, SeekOrigin.Begin);
                    result = outputStream.ToArray();
                }
            }
            return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
        }
    }
}