// Default URL for triggering event grid function in the local environment.
// http://localhost:7071/runtime/webhooks/EventGrid?functionName={functionname}

// Learn how to locally debug an Event Grid-triggered function:
//    https://aka.ms/AA30pjh

// Use for local testing:
//   https://{ID}.ngrok.io/runtime/webhooks/EventGrid?functionName=Thumbnail

using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.EventGrid.Models;
using Microsoft.Azure.WebJobs.Extensions.EventGrid;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;


using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace BlobPDFConverterFunction
{
    public static class BlobPDFConverter
    {
        private static readonly string BLOB_STORAGE_CONNECTION_STRING = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
        private static readonly string SOURCE_PATH_PREFIX = Environment.GetEnvironmentVariable("SOURCE_PATH_PREFIX");
        private static readonly string TARGET_PATH_PREFIX = Environment.GetEnvironmentVariable("TARGET_PATH_PREFIX");

        private static string GetTargetBlobNameFromUrl(string bloblUrl)
        {
            var uri = new Uri(bloblUrl);
            var cloudBlob = new CloudBlob(uri);
            return TARGET_PATH_PREFIX + "/" + cloudBlob.Name + ".pdf";
        }

        private static string GetTargetContainerNameFromUrl(string bloblUrl)
        {
            var uri = new Uri(bloblUrl);
            var cloudBlob = new CloudBlob(uri);
            return cloudBlob.Container.Name;
        }
        
        private static bool ShouldConvert(string url)
        {
            var extension = Path.GetExtension(url);
            extension = extension.Replace(".", "");
            return Regex.IsMatch(extension, "txt|docx|ppt", RegexOptions.IgnoreCase);
        }

        
        //setup ngrok: ngrok.com
        //https://e6d806f9.ngrok.io/runtime/webhooks/eventgrid?functionName=BlobConvertToPDFEventHandler

        // https://docs.microsoft.com/en-us/azure/azure-functions/functions-debug-event-grid-trigger-local
        //

        [FunctionName("BlobConvertToPDFEventHandler")]
        public static async Task Run(
            [EventGridTrigger]EventGridEvent eventGridEvent, 
            [Blob("{data.url}", FileAccess.Read)] Stream input,
            ILogger log)
        {
            try
            {
                if (input != null)
                {
                    var createdEvent = ((JObject)eventGridEvent.Data).ToObject<StorageBlobCreatedEventData>();
  
                    if (ShouldConvert(createdEvent.Url))
                    {
                        var storageAccount = CloudStorageAccount.Parse(BLOB_STORAGE_CONNECTION_STRING);
                        var blobClient = storageAccount.CreateCloudBlobClient();
                        
                        var targetContainer = blobClient.GetContainerReference(GetTargetContainerNameFromUrl(createdEvent.Url));
                        var targetBlobName = GetTargetBlobNameFromUrl(createdEvent.Url);
                        var targetBlockBlob = targetContainer.GetBlockBlobReference(targetBlobName);

                        using (var output = new MemoryStream())
                        using (WordDocument wordDocument = new WordDocument(input, FormatType.Automatic))
                        {
                            //Creates an instance of DocToPDFConverter - responsible for Word to PDF conversion
                            DocIORenderer converter = new DocIORenderer();
                            
                            //Converts Word document into PDF document
                            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
                            
                            //Save the document into stream.
                            pdfDocument.Save(output);

                            //Closes the instance of PDF document object
                            pdfDocument.Close();

                            output.Position = 0;
                            await targetBlockBlob.UploadFromStreamAsync(output);
                        }

                    }
                    else
                    {
                        log.LogInformation($"Will NOT convert: {createdEvent.Url}");
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogInformation(ex.Message);
                throw;
            }
        }
    }
}
