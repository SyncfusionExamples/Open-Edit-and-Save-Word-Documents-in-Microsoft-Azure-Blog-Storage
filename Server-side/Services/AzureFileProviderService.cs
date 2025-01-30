using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Specialized;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.FileManager.AzureFileProvider;
using Syncfusion.EJ2.FileManager.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace EJ2AzureASPCoreFileProvider.Services
{
    public interface IAzureFileProviderService
    {
        object AzureFileOperations(FileManagerDirectoryContent args);
        object Download(FileManagerDirectoryContent args);
        Task<IActionResult> GetDocumentAsync(Dictionary<string, string> jsonObject);
        Task SaveToAzureAsync(IFormCollection data);
        Task<IActionResult> DownloadFileAsync(string documentName);
        
    }

    public class AzureFileProviderService : IAzureFileProviderService
    {
        private readonly string _storageConnectionString;
        private readonly string _accountName;
        private readonly string _accountKey;
        private readonly string _containerName;
        private readonly ILogger<AzureFileProviderService> _logger;
        private readonly AzureFileProvider _fileProvider;

        public AzureFileProviderService(
            IConfiguration configuration,
            ILogger<AzureFileProviderService> logger)
        {
            _storageConnectionString = configuration["connectionString"];
            _accountName = configuration["accountName"];
            _accountKey = configuration["accountKey"];
            _containerName = configuration["containerName"];
            _logger = logger;

            // Initialize Syncfusion Azure File Provider
            _fileProvider = new AzureFileProvider();
            var basePath = "https://documenteditorstorage.blob.core.windows.net/inputfiles";
            var filePath = $"{basePath}/Files";
            _fileProvider.SetBlobContainer(basePath, filePath);
            _fileProvider.RegisterAzure(_accountName, _accountKey, _containerName);
        }

        public object PerformFileOperations(FileManagerDirectoryContent args)
        {
            try
            {
                NormalizePaths(ref args);

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
                _logger.LogError(ex, "File operation failed");
                throw;
            }
        }

        public object Download(FileManagerDirectoryContent args)
        {
            try
            {
                NormalizePaths(ref args); // Normalize paths before passing to the provider
                return _fileProvider.Download(args.Path, args.Names, args.Data);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Download operation failed");
                throw;
            }
        }
        public async Task<IActionResult> GetDocumentAsync(Dictionary<string, string> jsonObject)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                var documentName = jsonObject["documentName"];
                var blobPath = GetBlobPath(documentName);
                var blobClient = GetBlobClient(blobPath);

                if (await blobClient.ExistsAsync())
                {
                    await blobClient.DownloadToAsync(stream);
                    stream.Position = 0;
                    WordDocument document = WordDocument.Load(stream, FormatType.Docx);
                    string json = JsonConvert.SerializeObject(document);
                    document.Dispose();

                    return new OkObjectResult(json);
                }
                return new NotFoundResult();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Document retrieval failed");
                return new StatusCodeResult(500);
            }
            finally
            {
                stream.Dispose();
            }
        }

        public async Task SaveToAzureAsync(IFormCollection data)
        {
            try
            {
                var file = data.Files[0];
                var documentName = GetValue(data, "documentName");
                var blobPath = GetBlobPath(documentName);
                var blobClient = GetBlobClient(blobPath);

                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                await blobClient.UploadAsync(stream);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "File upload failed");
                throw;
            }
        }

        public async Task<IActionResult> DownloadFileAsync(string documentName)
        {
            try
            {
                var blobPath = GetBlobPath(documentName);
                var blobClient = GetBlobClient(blobPath);

                if (!await blobClient.ExistsAsync())
                    return new NotFoundResult();

                var stream = new MemoryStream();
                await blobClient.DownloadToAsync(stream);
                stream.Position = 0;

                return new FileStreamResult(stream, "application/octet-stream")
                {
                    FileDownloadName = documentName
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "File download failed");
                return new StatusCodeResult(500);
            }
        }

        private void NormalizePaths(ref FileManagerDirectoryContent args)
        {
            if (string.IsNullOrEmpty(args.Path)) return;

            const string basePath = "https://documenteditorstorage.blob.core.windows.net/inputfiles/";
            var originalPath = $"{basePath}Files".Replace(basePath, "");

            args.Path = args.Path.Contains(originalPath)
                ? args.Path.Replace("//", "/")
                : $"{originalPath}{args.Path}".Replace("//", "/");

            args.TargetPath = $"{originalPath}{args.TargetPath}".Replace("//", "/");
        }

        private BlockBlobClient GetBlobClient(string blobPath)
        {
            var serviceClient = new BlobServiceClient(_storageConnectionString);
            var containerClient = serviceClient.GetBlobContainerClient(_containerName);
            return containerClient.GetBlockBlobClient(blobPath);
        }

        private string GetBlobPath(string documentName) => $"Files/{documentName}";

        private static string GetValue(IFormCollection data, string key) =>
            data.TryGetValue(key, out var values) && values.Count > 0
                ? values[0]
                : string.Empty;
    }
}