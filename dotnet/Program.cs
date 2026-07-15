using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;

var configuration = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: false)
    .Build();

var tenantId = configuration["AzureAd:TenantId"];
var siteHostname = configuration["SharePoint:SiteHostname"]!;
var sitePath = configuration["SharePoint:SitePath"]!;
var driveId = configuration["SharePoint:DriveId"];
var folderPath = configuration["SharePoint:FolderPath"] ?? "";

if (args.Length == 0)
{
    Console.WriteLine("Usage: sharepoint-upload-documents <file-path> [destination-filename]");
    return 1;
}

var filePath = args[0];
if (!File.Exists(filePath))
{
    Console.WriteLine($"Error: File not found: {filePath}");
    return 1;
}

var destinationFileName = args.Length > 1 ? args[1] : Path.GetFileName(filePath);

// Authenticate using DefaultAzureCredential (supports managed identity, az login, env vars, etc.)
var credentialOptions = new DefaultAzureCredentialOptions();
if (!string.IsNullOrEmpty(tenantId))
    credentialOptions.TenantId = tenantId;
var credential = new DefaultAzureCredential(credentialOptions);
var graphClient = new GraphServiceClient(credential);

try
{
    // Resolve the SharePoint site
    Console.WriteLine($"Resolving site: {siteHostname}{sitePath}");
    var site = await graphClient.Sites[$"{siteHostname}:{sitePath}"].GetAsync();
    if (site == null)
    {
        Console.WriteLine("Error: Could not resolve SharePoint site.");
        return 1;
    }

    // Get the document library drive
    string targetDriveId;
    if (!string.IsNullOrEmpty(driveId))
    {
        targetDriveId = driveId;
    }
    else
    {
        // Use the default document library
        var drive = await graphClient.Sites[site.Id].Drive.GetAsync();
        if (drive == null)
        {
            Console.WriteLine("Error: Could not resolve default document library.");
            return 1;
        }
        targetDriveId = drive.Id!;
    }

    var fileInfo = new FileInfo(filePath);
    var itemPath = string.IsNullOrEmpty(folderPath)
        ? destinationFileName
        : $"{folderPath.TrimEnd('/')}/{destinationFileName}";

    Console.WriteLine($"Uploading '{fileInfo.Name}' ({fileInfo.Length / 1024.0:F1} KB) to: {itemPath}");

    string? uploadedItemId = null;
    string? uploadedItemUrl = null;

    if (fileInfo.Length <= 4 * 1024 * 1024) // 4 MB threshold
    {
        // Simple upload for small files
        await using var stream = File.OpenRead(filePath);
        var uploadedItem = await graphClient.Drives[targetDriveId]
            .Items["root"]
            .ItemWithPath(itemPath)
            .Content
            .PutAsync(stream);

        uploadedItemId = uploadedItem?.Id;
        uploadedItemUrl = uploadedItem?.WebUrl;
    }
    else
    {
        // Resumable upload session for large files (> 4 MB)
        var uploadSessionRequest = new CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    { "@microsoft.graph.conflictBehavior", "replace" }
                }
            }
        };

        var uploadSession = await graphClient.Drives[targetDriveId]
            .Items["root"]
            .ItemWithPath(itemPath)
            .CreateUploadSession
            .PostAsync(uploadSessionRequest);

        if (uploadSession == null)
        {
            Console.WriteLine("Error: Failed to create upload session.");
            return 1;
        }

        // Max slice size must be a multiple of 320 KiB (327,680 bytes)
        const int maxSliceSize = 10 * 320 * 1024; // ~3.2 MB per slice

        await using var fileStream = File.OpenRead(filePath);
        var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

        var totalLength = fileStream.Length;
        var progress = new Progress<long>(uploaded =>
        {
            var percent = (double)uploaded / totalLength * 100;
            Console.Write($"\rProgress: {percent:F1}%   ");
        });

        var uploadResult = await fileUploadTask.UploadAsync(progress);

        Console.WriteLine();
        if (uploadResult.UploadSucceeded)
        {
            uploadedItemId = uploadResult.ItemResponse?.Id;
            uploadedItemUrl = uploadResult.ItemResponse?.WebUrl;
        }
        else
        {
            Console.WriteLine("Error: Upload failed.");
            return 1;
        }
    }

    Console.WriteLine($"Upload complete. Item ID: {uploadedItemId}, url: {uploadedItemUrl}");

    // Convert to PDF if office document
    var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();
    if (uploadedItemId != null && fileExtension is ".docx" or ".doc" or ".xlsx" or ".xls" or ".pptx" or ".ppt")
    {
        Console.WriteLine("Converting to PDF...");
        var pdfStream = await graphClient.Drives[targetDriveId]
            .Items[uploadedItemId]
            .Content
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Format = "pdf";
            });

        var pdfFileName = Path.ChangeExtension(destinationFileName, ".pdf");
        var pdfItemPath = string.IsNullOrEmpty(folderPath)
            ? pdfFileName
            : $"{folderPath.TrimEnd('/')}/{pdfFileName}";

        await using var pdfUploadStream = new MemoryStream();
        await pdfStream!.CopyToAsync(pdfUploadStream);
        pdfUploadStream.Position = 0;

        var uploadedPdfItem = await graphClient.Drives[targetDriveId]
            .Items["root"]
            .ItemWithPath(pdfItemPath)
            .Content
            .PutAsync(pdfUploadStream);

        Console.WriteLine($"PDF conversion complete. Item ID: {uploadedPdfItem?.Id}, url: {uploadedPdfItem?.WebUrl}");
    }
    // end of pdf conversion section
    
}
catch (ServiceException ex)
{
    Console.WriteLine($"Graph API error: {ex.Message}");
    Console.WriteLine($"Status code: {ex.ResponseStatusCode}");
    return 1;
}

return 0;
