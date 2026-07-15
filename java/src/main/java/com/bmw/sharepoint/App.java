package com.bmw.sharepoint;

import com.azure.core.credential.AccessToken;
import com.azure.core.credential.TokenCredential;
import com.azure.core.credential.TokenRequestContext;
import com.microsoft.graph.serviceclient.GraphServiceClient;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.OffsetDateTime;
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

        // Authenticate using az CLI with explicit tenant
        TokenCredential credential = createCliCredential(config.tenantId());
        var graphClient = new GraphServiceClient(credential, "https://graph.microsoft.com/.default");

        try {
            // Resolve the SharePoint site ID via direct HTTP (Kiota encodes colons in path-based addressing)
            logger.info("Resolving site: " + config.siteHostname() + config.sitePath());
            String siteId = resolveSiteId(credential, config.tenantId(), config.siteHostname(), config.sitePath());
            if (siteId == null) {
                logger.severe("Could not resolve SharePoint site.");
                return 1;
            }

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

    private static String resolveSiteId(com.azure.core.credential.TokenCredential credential, String tenantId, String hostname, String sitePath) {
        try {
            var tokenRequest = new com.azure.core.credential.TokenRequestContext()
                    .addScopes("https://graph.microsoft.com/.default");
            if (tenantId != null) {
                tokenRequest.setTenantId(tenantId);
            }
            String token = credential.getTokenSync(tokenRequest).getToken();

            var httpClient = java.net.http.HttpClient.newHttpClient();
            var uri = java.net.URI.create("https://graph.microsoft.com/v1.0/sites/" + hostname + ":" + sitePath);
            var request = java.net.http.HttpRequest.newBuilder(uri)
                    .header("Authorization", "Bearer " + token)
                    .GET()
                    .build();

            var response = httpClient.send(request, java.net.http.HttpResponse.BodyHandlers.ofString());
            if (response.statusCode() != 200) {
                logger.severe("Site resolution failed (HTTP " + response.statusCode() + "): " + response.body());
                return null;
            }

            // Parse "id" from JSON response (avoid adding JSON dependency just for this)
            String body = response.body();
            int idIdx = body.indexOf("\"id\"");
            if (idIdx < 0) return null;
            int colonIdx = body.indexOf(':', idIdx);
            int startQuote = body.indexOf('"', colonIdx + 1);
            int endQuote = body.indexOf('"', startQuote + 1);
            return body.substring(startQuote + 1, endQuote);
        } catch (Exception e) {
            logger.log(Level.SEVERE, "Failed to resolve site ID", e);
            return null;
        }
    }

    private static TokenCredential createCliCredential(String tenantId) {
        return new TokenCredential() {
            private AccessToken cachedToken;

            @Override
            public reactor.core.publisher.Mono<AccessToken> getToken(TokenRequestContext context) {
                return reactor.core.publisher.Mono.fromCallable(() -> getTokenSync(context));
            }

            @Override
            public synchronized AccessToken getTokenSync(TokenRequestContext context) {
                if (cachedToken != null && cachedToken.getExpiresAt().isAfter(OffsetDateTime.now().plusMinutes(5))) {
                    return cachedToken;
                }
                try {
                    var command = new java.util.ArrayList<String>();
                    boolean isWindows = System.getProperty("os.name").toLowerCase().contains("win");
                    if (isWindows) {
                        command.add("cmd");
                        command.add("/c");
                    }
                    command.add("az");
                    command.add("account");
                    command.add("get-access-token");
                    command.add("--resource");
                    command.add("https://graph.microsoft.com");
                    if (tenantId != null) {
                        command.add("--tenant");
                        command.add(tenantId);
                    }
                    command.add("--query");
                    command.add("accessToken");
                    command.add("-o");
                    command.add("tsv");

                    var process = new ProcessBuilder(command)
                            .redirectErrorStream(true)
                            .start();
                    String token = new String(process.getInputStream().readAllBytes()).trim();
                    int exitCode = process.waitFor();
                    if (exitCode != 0) {
                        throw new RuntimeException("az CLI failed (exit " + exitCode + "): " + token);
                    }
                    cachedToken = new AccessToken(token, OffsetDateTime.now().plusHours(1));
                    return cachedToken;
                } catch (Exception e) {
                    throw new RuntimeException("Failed to get token via az CLI", e);
                }
            }
        };
    }
}
