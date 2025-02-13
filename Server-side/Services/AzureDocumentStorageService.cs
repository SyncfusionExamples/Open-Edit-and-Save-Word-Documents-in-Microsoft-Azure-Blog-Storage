using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Specialized;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.FileManager.AzureDocumentManager;
using Syncfusion.EJ2.FileManager.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace EJ2AzureASPCoreFileProvider.Services
{
    public interface IAzureDocumentStorageService
    {
        object ManageDocument(FileManagerDirectoryContent args);
        object DownloadDocument(FileManagerDirectoryContent args);
        Task<IActionResult> FetchDocumentAsync(Dictionary<string, string> jsonObject);
        Task UploadDocumentAsync(IFormCollection data);
        Task<bool> CheckDocumentExistsAsync(string documentName);

    }

    /// <summary>
    /// Service for handling Azure storage operations using Syncfusion components
    /// </summary>
    public class AzureDocumentStorageService : IAzureDocumentStorageService
    {
        private readonly string _storageConnectionString;
        private readonly string _accountName;
        private readonly string _accountKey;
        private readonly string _containerName;
        private readonly ILogger<AzureDocumentStorageService> _logger;
        private readonly AzureDocumentManager _fileProvider;

        /// <summary>
        /// Initializes Azure storage configuration and file provider
        /// </summary>
        /// <param name="configuration">Application configuration settings</param>
        /// <param name="logger">Logger instance for error tracking</param>
        public AzureDocumentStorageService(
            IConfiguration configuration,
            ILogger<AzureDocumentStorageService> logger)
        {
            // Retrieve necessary configuration values for connecting to Azure storage.
            _storageConnectionString = configuration["connectionString"];
            _accountName = configuration["accountName"];
            _accountKey = configuration["accountKey"];
            _containerName = configuration["containerName"];
            _logger = logger;

            // Initialize Syncfusion Azure File Provider instance.
            _fileProvider = new AzureDocumentManager();

            // Define the base path and file path for the blob storage.
            var basePath = $"https://documenteditorstorage.blob.core.windows.net/{_containerName}";
            var filePath = $"{basePath}/Files";

            // Set the base blob container path for the file provider.
            _fileProvider.SetBlobContainer(basePath, filePath);
            // Register the Azure storage credentials and container name.
            _fileProvider.RegisterAzure(_accountName, _accountKey, _containerName);

            //----------
            //For example 
            //_fileProvider.setBlobContainer("https://azure_service_account.blob.core.windows.net/{containerName}/", "https://azure_service_account.blob.core.windows.net/{containerName}/Files");
            //_fileProvider.RegisterAzure("azure_service_account", "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", "containerName");
            // Note: we need to create a Files folder inside the container and the documents folder inside the folder can be accessed.
            //---------
        }

        /// <summary>
        /// Executes file management operations against Azure storage
        /// </summary>
        /// <param name="args">Operation parameters including action type and paths</param>
        /// <returns>Operation result in camelCase format</returns>
        /// <exception cref="Exception">Thrown for Azure storage operation failures</exception>
        public object ManageDocument(FileManagerDirectoryContent args)
        {
            try
            {
                // Normalize the incoming paths to ensure they are in the expected format.
                NormalizeDocumentPaths(ref args);
                // Determine the action and execute the corresponding method on the file provider.
                return args.Action switch
                {
                    "read" => _fileProvider.ToCamelCase(_fileProvider.GetFiles(args.Path, args.ShowHiddenItems, args.Data)),
                    "delete" => _fileProvider.ToCamelCase(_fileProvider.Delete(args.Path, args.Names, args.Data)),
                    "details" => _fileProvider.ToCamelCase(_fileProvider.Details(args.Path, args.Names, args.Data)),
                    "search" => _fileProvider.ToCamelCase(_fileProvider.Search(
                        args.Path, args.SearchString, args.ShowHiddenItems, args.CaseSensitive, args.Data)),
                    "copy" => _fileProvider.ToCamelCase(_fileProvider.Copy(
                        args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData, args.Data)),
                    _ => null
                };
            }
            catch (Exception ex)
            {
                // Log any errors encountered during file operations.
                _logger.LogError(ex, "File operation failed");
                throw;
            }
        }

        /// <summary>
        /// Retrieves a Word document from Azure and converts it to JSON format
        /// </summary>
        /// <param name="jsonObject">Contains document name for lookup</param>
        /// <returns>Word document content in JSON format</returns>
        public async Task<IActionResult> FetchDocumentAsync(Dictionary<string, string> jsonObject)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                // Extract the document name from the provided JSON object.
                var documentName = jsonObject["documentName"];
                // Build the blob path for the document.
                var blobPath = GenerateDocumentBlobPath(documentName);
                // Get a reference to the blob client for the specified document.
                var blobClient = CreateBlobClient(blobPath);
                // Check if the blob exists in the container.
                if (await blobClient.ExistsAsync())
                {
                    // Download the blob content into the memory stream.
                    await blobClient.DownloadToAsync(stream);
                    stream.Position = 0;
                    // Load the WordDocument from the stream.
                    WordDocument document = WordDocument.Load(stream, FormatType.Docx);
                    // Serialize the document to JSON using Newtonsoft.Json.
                    string json = JsonConvert.SerializeObject(document);
                    // Dispose of the document after serialization.
                    document.Dispose();
                    // Return the JSON content with an OK (200) status.
                    return new OkObjectResult(json);
                }
                // If the blob doesn't exist, return a 404 Not Found response.
                return new NotFoundResult();
            }
            catch (Exception ex)
            {
                // Log any exceptions and return a 500 Internal Server Error.
                _logger.LogError(ex, "Document retrieval failed");
                return new StatusCodeResult(500);
            }
            finally
            {
                stream.Dispose();
            }
        }

        /// <summary>
        /// Handles file download operations from Azure file manager
        /// </summary>
        /// <param name="args">Download parameters including target paths</param>
        /// <returns>File stream result or error response</returns>
        /// <exception cref="Exception">Thrown for download failures</exception>
        public object DownloadDocument(FileManagerDirectoryContent args)
        {
            try
            {
                // Normalize paths before processing the download operation.
                NormalizeDocumentPaths(ref args);
                // Delegate the download operation to the file provider.
                return _fileProvider.Download(args.Path, args.Names, args.Data);
            }
            catch (Exception ex)
            {
                // Log any errors that occur during the download process.
                _logger.LogError(ex, "Download operation failed");
                throw;
            }
        }

        /// <summary>
        /// Saves a document file to Azure storage
        /// </summary>
        /// <param name="data">Form data containing the file to save</param>
        /// <exception cref="Exception">Thrown for save failures</exception>
        public async Task UploadDocumentAsync(IFormCollection data)
        {
            try
            {
                // Retrieve the first file from the form data.
                var file = data.Files[0];
                // Get the document name from the form collection.
                var documentName = ExtractFormValue(data, "documentName");
                // Construct the blob path based on the document name.
                var blobPath = GenerateDocumentBlobPath(documentName);
                // Check if the blob already exists.
                var blobClient = CreateBlobClient(blobPath);
                if(blobClient.Exists())
                {
                    // Upload the file content to the existing blob.
                    using var stream = new MemoryStream();
                    await file.CopyToAsync(stream);
                    stream.Position = 0;
                    await blobClient.UploadAsync(stream);
                }
                else
                {
                    // If the blob does not exist, uploading an empty stream 
                    using var stream = new MemoryStream();
                    await blobClient.UploadAsync(stream);
                }
               
            }
            catch (Exception ex)
            {
                // Log errors during file upload and rethrow the exception.
                _logger.LogError(ex, "File upload failed");
                throw;
            }
        }

        /// <summary>
        /// Normalizes Azure blob paths for Syncfusion file provider compatibility
        /// </summary>
        /// <param name="args">FileManagerDirectoryContent reference to modify</param>
        private void NormalizeDocumentPaths(ref FileManagerDirectoryContent args)
        {
            if (string.IsNullOrEmpty(args.Path)) return;

            // Define the base path used in the blob storage URL.
            var basePath = $"https://documenteditorstorage.blob.core.windows.net/{_containerName}/";
            var originalPath = $"{basePath}Files".Replace(basePath, "");

            args.Path = args.Path.Contains(originalPath)
                ? args.Path.Replace("//", "/")
                : $"{originalPath}{args.Path}".Replace("//", "/");

            args.TargetPath = $"{originalPath}{args.TargetPath}".Replace("//", "/");
        }

        // Add to AzureFileProviderService class
        public async Task<bool> CheckDocumentExistsAsync(string documentName)
        {
            // Construct the blob path for the document based on the document name.
            var blobPath = GenerateDocumentBlobPath(documentName);
            // Get the BlockBlobClient for the specified blob path.
            var blobClient = CreateBlobClient(blobPath);
            // Use the Azure Blob SDK to check if the blob exists.
            return await blobClient.ExistsAsync();
        }

        /// <summary>
        /// Creates and returns a BlockBlobClient for interacting with a specific blob
        /// </summary>
        /// <param name="blobPath">The full path to the blob within the container</param>
        /// <returns>Configured BlockBlobClient instance</returns>
        private BlockBlobClient CreateBlobClient(string blobPath)
        {
            var serviceClient = new BlobServiceClient(_storageConnectionString);
            var containerClient = serviceClient.GetBlobContainerClient(_containerName);
            return containerClient.GetBlockBlobClient(blobPath);
        }

        /// <summary>
        /// Generates the full blob path by combining the document name with the base file path
        /// </summary>
        /// <param name="documentName">Name of the target document</param>
        /// <returns>Full blob path in format 'Files/{documentName}'</returns>
        private string GenerateDocumentBlobPath(string documentName) => $"Files/{documentName}";

        /// <summary>
        /// Safely retrieves a value from form collection data
        /// </summary>
        /// <param name="data">Form collection containing request data</param>
        /// <param name="key">Key to look up in the form data</param>
        /// <returns>First value for the key if exists, otherwise empty string</returns>
        private static string ExtractFormValue(IFormCollection data, string key) =>
            data.TryGetValue(key, out var values) && values.Count > 0
                ? values[0]
                : string.Empty;
    }
}