package com.bmw.sharepoint;

import com.azure.core.credential.TokenCredential;
import com.azure.identity.DefaultAzureCredentialBuilder;
import com.microsoft.graph.serviceclient.GraphServiceClient;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

public class App {

    private static final Logger logger = Logger.getLogger(App.class.getName());

    record AppConfig(
            String tenantId,
            String siteHostname,
            String sitePath,
            String driveId,
            String folderPath
    ) {
        static AppConfig load() {
            var props = new Properties();
            try (var is = App.class.getResourceAsStream("/application.properties")) {
                if (is != null) {
                    props.load(is);
                }
            } catch (IOException e) {
                logger.warning("Could not load application.properties from classpath: " + e.getMessage());
            }

            return new AppConfig(
                    resolve(props, "azure.tenant-id", "SHAREPOINT_AZURE_TENANT_ID"),
                    resolve(props, "sharepoint.site-hostname", "SHAREPOINT_SITE_HOSTNAME"),
                    resolve(props, "sharepoint.site-path", "SHAREPOINT_SITE_PATH"),
                    resolve(props, "sharepoint.drive-id", "SHAREPOINT_DRIVE_ID"),
                    resolve(props, "sharepoint.folder-path", "SHAREPOINT_FOLDER_PATH")
            );
        }

        private static String resolve(Properties props, String key, String envVar) {
            String envValue = System.getenv(envVar);
            if (envValue != null && !envValue.isBlank()) {
                return envValue;
            }
            String propValue = props.getProperty(key, "");
            return propValue.isBlank() ? null : propValue;
        }
    }

    public static void main(String[] args) {
        System.exit(new App().run(args));
    }

    private int run(String[] args) {
        var config = AppConfig.load();

        if (args.length == 0) {
            System.out.println("Usage: sharepoint-upload-documents <file-path> [destination-filename]");
            return 1;
        }

        Path filePath = Path.of(args[0]);
        if (!Files.exists(filePath)) {
            logger.severe("File not found: " + filePath);
            return 1;
        }

        String destinationFileName = args.length > 1 ? args[1] : filePath.getFileName().toString();
        String folderPath = config.folderPath() != null ? config.folderPath() : "";

        // Authenticate using DefaultAzureCredential (supports az CLI, managed identity, env vars, etc.)
        var credentialBuilder = new DefaultAzureCredentialBuilder();
        if (config.tenantId() != null) {
            credentialBuilder.tenantId(config.tenantId());
        }
        TokenCredential credential = credentialBuilder.build();
        var graphClient = new GraphServiceClient(credential, "https://graph.microsoft.com/.default");

        try {
            // Resolve the SharePoint site
            logger.info("Resolving site: " + config.siteHostname() + config.sitePath());
            var site = graphClient.sites().bySiteId(config.siteHostname() + ":" + config.sitePath() + ":").get();
            if (site == null) {
                logger.severe("Could not resolve SharePoint site.");
                return 1;
            }
            String siteId = site.getId();

            // Get the document library drive
            String targetDriveId;
            if (config.driveId() != null) {
                targetDriveId = config.driveId();
            } else {
                var drive = graphClient.sites().bySiteId(siteId).drive().get();
                if (drive == null) {
                    logger.severe("Could not resolve default document library.");
                    return 1;
                }
                targetDriveId = drive.getId();
            }

            // Build remote path
            String remotePath = folderPath.isEmpty()
                    ? destinationFileName
                    : folderPath.replaceAll("/+$", "") + "/" + destinationFileName;

            // Upload the file
            var uploader = new SharePointUploader(graphClient, targetDriveId);
            var uploadResult = uploader.upload(filePath, remotePath);
            if (uploadResult == null) {
                return 1;
            }

            // Convert to PDF if it's an Office document
            var pdfConverter = new PdfConverter(graphClient, targetDriveId);
            if (pdfConverter.isConvertible(destinationFileName)) {
                String pdfFileName = changeExtension(destinationFileName, ".pdf");
                String pdfRemotePath = folderPath.isEmpty()
                        ? pdfFileName
                        : folderPath.replaceAll("/+$", "") + "/" + pdfFileName;
                pdfConverter.convertAndUpload(uploadResult.itemId(), pdfRemotePath);
            }

        } catch (com.microsoft.graph.models.odataerrors.ODataError ex) {
            logger.log(Level.SEVERE, "Graph API error (HTTP " + ex.getResponseStatusCode() + ")", ex);
            return 1;
        } catch (Exception ex) {
            logger.log(Level.SEVERE, "Unexpected error", ex);
            return 1;
        }

        return 0;
    }

    private static String changeExtension(String fileName, String newExtension) {
        int lastDot = fileName.lastIndexOf('.');
        String baseName = lastDot >= 0 ? fileName.substring(0, lastDot) : fileName;
        return baseName + newExtension;
    }


}
