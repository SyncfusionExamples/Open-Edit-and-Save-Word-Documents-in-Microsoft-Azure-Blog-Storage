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
    /// <summary>
    /// Controller for handling Azure file operations and document management
    /// </summary>
    [Route("api/[controller]")]
    [EnableCors("AllowAllOrigins")]
    public class AzureFileProviderController : ControllerBase
    {
        private readonly IAzureFileProviderService _fileService;

        /// <summary>
        /// Constructor injecting the file provider service dependency.
        /// </summary>
        /// <param name="fileService">Service for performing file operations</param>
        public AzureFileProviderController(IAzureFileProviderService fileService)
        {
            _fileService = fileService;
        }

        /// <summary>
        /// Handles file management operations (read, delete, copy, search)
        /// </summary>
        /// <param name="args">File operation parameters including path and action type</param>
        /// <returns>Result of the file operation</returns>
        [HttpPost("AzureFileOperations")]
        public object AzureFileOperations([FromBody] FileManagerDirectoryContent args)
        {
            return _fileService.PerformFileOperations(args);
        }

        /// <summary>
        /// Downloads selected files or folders from Azure file manager
        /// </summary>
        /// <param name="downloadInput">JSON string containing download parameters</param>
        /// <returns>File content stream or error response</returns>
        [HttpPost("AzureDownload")]
        public object AzureDownload(string downloadInput)
        {
            if(downloadInput !=null)
            {
                // Set serializer options to use camelCase naming policy.
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                };
                // Deserialize the JSON string to a FileManagerDirectoryContent object
                FileManagerDirectoryContent args = JsonSerializer.Deserialize<FileManagerDirectoryContent>(downloadInput, options);
                return _fileService.Download(args);
            }
            // Return null if input is not provided
            return null;
        }

        /// <summary>
        /// Retrieves a document from Azure storage in JSON format
        /// </summary>
        /// <param name="jsonObject">Contains document name and metadata</param>
        /// <returns>Document content as JSON or error response</returns>
        [HttpPost("GetFile")]
        public async Task<IActionResult> GetFile([FromBody] Dictionary<string, string> jsonObject)
        {
            return await _fileService.GetDocumentAsync(jsonObject);
        }

        /// <summary>
        /// Saves uploaded document to Azure storage
        /// </summary>
        /// <param name="data">Form data containing file and document name</param>
        [HttpPost("SaveToAzure")]
        public async Task SaveToAzure(IFormCollection data)
        {
            await _fileService.SaveToAzureAsync(data);
        }

        /// <summary>
        /// Checks if a document with the given name exists in the Azure Storage container.
        /// Expects a JSON payload with a "fileName" property.
        /// </summary>
        /// <param name="jsonObject">
        /// A dictionary containing the document name to check. For example: { "fileName": "Document1.docx" }.
        /// </param>
        /// <returns>
        /// An <see cref="IActionResult"/> containing a JSON object with a boolean property "exists".
        /// If the document exists, the response will be { "exists": true }; otherwise, { "exists": false }.
        /// </returns>
        [HttpPost("ValidateFileExistence")]
        public async Task<IActionResult> ValidateFileExistence([FromBody] Dictionary<string, string> jsonObject)
        {
            // Validate that the "fileName" key exists in the request payload.
            if (!jsonObject.TryGetValue("fileName", out var fileName) || string.IsNullOrEmpty(fileName))
            {
                return BadRequest("fileName not provided");
            }

            try
            {
                // Call the service method to check if the document exists.
                bool exists = await _fileService.CheckDocumentExistsAsync(fileName);
                // Return a 200 OK response with the result.
                return Ok(new { exists });
            }
            catch (Exception ex)
            {
                // If an error occurs, return a 500 Internal Server Error with the error message.
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Helper class for clipboard operation parameters
        /// </summary>
        public class CustomParameter
        {
            /// <summary>
            /// Document content in base64 string format
            /// </summary>
            public string content { get; set; }

            /// <summary>
            /// File extension type (e.g., .docx)
            /// </summary>
            public string type { get; set; }
        }

        /// <summary>
        /// Processes clipboard content for document editor compatibility
        /// </summary>
        /// <param name="param">Clipboard content and type information</param>
        /// <returns>Serialized document JSON or empty string on error</returns>
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody] CustomParameter param)
        {
            // Check if the clipboard content is not null or empty.
            if (param.content != null && param.content != "")
            {
                try
                {
                    // Load the WordDocument from the provided content using the appropriate format.
                    WordDocument document = WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                    // Serialize the WordDocument to JSON format using Newtonsoft.Json.
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                    // Dispose of the document to free resources.
                    document.Dispose();
                    return json;
                }
                catch (Exception)
                {
                    // Return empty string on any exception
                    return "";
                }
            }
            // Return empty string if content is null or empty.
            return "";
        }

        /// <summary>
        /// Converts file extension to DocumentEditor format type
        /// </summary>
        /// <param name="format">File extension (e.g., ".docx", ".txt")</param>
        /// <returns>Corresponding FormatType enum value</returns>
        /// <exception cref="NotSupportedException">Thrown for unsupported file formats</exception>
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