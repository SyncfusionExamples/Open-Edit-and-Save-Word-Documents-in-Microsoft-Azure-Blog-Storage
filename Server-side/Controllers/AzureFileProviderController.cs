using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.EJ2.FileManager.Base;
using System.Collections.Generic;
using System.Threading.Tasks;
using EJ2AzureASPCoreFileProvider.Services;
using Microsoft.AspNetCore.Http;
using Azure;
using System.Text.Json;
using System;
using Syncfusion.EJ2.DocumentEditor;

namespace EJ2AzureASPCoreFileProvider.Controllers
{
    [Route("api/[controller]")]
    [EnableCors("AllowAllOrigins")]
    public class AzureFileProviderController : ControllerBase
    {
        private readonly IAzureFileProviderService _fileService;

        public AzureFileProviderController(IAzureFileProviderService fileService)
        {
            _fileService = fileService;
        }

        [HttpPost("AzureFileOperations")]
        public object AzureFileOperations([FromBody] FileManagerDirectoryContent args)
        {
            return _fileService.PerformFileOperations(args);
        }

        // Download the selected file(s) and folder(s)
        [HttpPost("AzureDownload")]
        public object AzureDownload(string downloadInput)
        {
            if(downloadInput !=null)
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                };
                FileManagerDirectoryContent args = JsonSerializer.Deserialize<FileManagerDirectoryContent>(downloadInput, options);
                return _fileService.Download(args);
            }
            
            return null;
        }

        [HttpPost("GetDocument")]
        public async Task<IActionResult> GetDocument([FromBody] Dictionary<string, string> jsonObject)
        {
            return await _fileService.GetDocumentAsync(jsonObject);
        }

        [HttpPost("SaveToAzure")]
        public async Task SaveToAzure(IFormCollection data)
        {
            await _fileService.SaveToAzureAsync(data);
        }

        [HttpGet("DownloadFile")]
        public async Task<IActionResult> DownloadFile([FromQuery] string documentName)
        {
            return await _fileService.DownloadFileAsync(documentName);
        }

        public class CustomParameter
        {
            public string content { get; set; }
            public string type { get; set; }
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody] CustomParameter param)
        {
            if (param.content != null && param.content != "")
            {
                try
                {
                    WordDocument document = WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                    document.Dispose();
                    return json;
                }
                catch (Exception)
                {
                    return "";
                }
            }
            return "";
        }

        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }
    }
}