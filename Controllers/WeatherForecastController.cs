using Microsoft.AspNetCore.Mvc;
using System.Text;
using Aspose.Words;
using SautinSoft.Document;
using SautinSoft;
using System;

namespace HtmlToDocx.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
     
        public WeatherForecastController()
        {
        }

        [HttpPost]
        [Route("/SautinSoft")]
        public IActionResult ConvertHtmlToDocx([FromBody] string html)
        {
            DocumentCore dc = DocumentCore.Load(new MemoryStream(Encoding.UTF8.GetBytes(html)), new HtmlLoadOptions());

            MemoryStream outputStream = new MemoryStream();
            dc.Save(outputStream, new DocxSaveOptions());
            outputStream.Position = 0;

            return File(outputStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
        }

        [HttpPost]
        [Route("/SautinSoft/2")]
        public IActionResult ConvertHtmlToDocx2([FromBody] string html)
        {
            HtmlToRtf h = new HtmlToRtf();
            byte[] byteArray = Encoding.UTF8.GetBytes(html);
            MemoryStream stream = new MemoryStream(byteArray);
            h.OpenHtml(stream);
            MemoryStream outputStream = new MemoryStream();
            h.ToDocx(outputStream);
            outputStream.Position = 0;
            return File(outputStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Output.docx");
        }

        [HttpPost]
        [Route("/Aspose")]
        public IActionResult ConvertHtmlToDocxAspose([FromBody] string html)
        {
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)));
            var memory = new MemoryStream();
            doc.Save(memory, SaveFormat.Docx);
            memory.Position = 0;
            return File(memory, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
        }
    }
}