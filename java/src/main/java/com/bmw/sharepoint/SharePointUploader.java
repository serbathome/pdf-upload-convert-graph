package com.bmw.sharepoint;

import com.microsoft.graph.core.models.IProgressCallback;
import com.microsoft.graph.core.tasks.LargeFileUploadTask;
import com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionPostRequestBody;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.DriveItemUploadableProperties;
import com.microsoft.graph.models.UploadSession;
import com.microsoft.graph.serviceclient.GraphServiceClient;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.logging.Logger;

public class SharePointUploader {

    private static final Logger logger = Logger.getLogger(SharePointUploader.class.getName());
    private static final long SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024; // 4 MB
    private static final int SLICE_SIZE = 10 * 320 * 1024; // ~3.2 MB, must be multiple of 320 KiB

    private final GraphServiceClient graphClient;
    private final String driveId;

    public SharePointUploader(GraphServiceClient graphClient, String driveId) {
        this.graphClient = graphClient;
        this.driveId = driveId;
    }

    public record UploadResult(String itemId, String webUrl) {}

    public UploadResult upload(Path localFile, String remotePath) throws IOException, ReflectiveOperationException, InterruptedException {
        long fileSize = Files.size(localFile);
        String itemId = "root:/" + remotePath + ":";

        logger.info("Uploading '%s' (%.1f KB) to: %s".formatted(
                localFile.getFileName(), fileSize / 1024.0, remotePath));

        DriveItem uploaded = (fileSize <= SIMPLE_UPLOAD_LIMIT)
                ? uploadSimple(localFile, itemId)
                : uploadLarge(localFile, itemId, fileSize);

        if (uploaded == null) {
            return null;
        }

        logger.info("Upload complete. Item ID: " + uploaded.getId() + ", url: " + uploaded.getWebUrl());
        return new UploadResult(uploaded.getId(), uploaded.getWebUrl());
    }

    private DriveItem uploadSimple(Path localFile, String itemId) throws IOException {
        try (InputStream stream = Files.newInputStream(localFile)) {
            return graphClient.drives().byDriveId(driveId)
                    .items().byDriveItemId(itemId)
                    .content()
                    .put(stream);
        }
    }

    private DriveItem uploadLarge(Path localFile, String itemId, long fileSize)
            throws IOException, ReflectiveOperationException, InterruptedException {

        var uploadableProperties = new DriveItemUploadableProperties();
        var additionalData = new HashMap<String, Object>();
        additionalData.put("@microsoft.graph.conflictBehavior", "replace");
        uploadableProperties.setAdditionalData(additionalData);

        var requestBody = new CreateUploadSessionPostRequestBody();
        requestBody.setItem(uploadableProperties);

        UploadSession session = graphClient.drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .createUploadSession()
                .post(requestBody);

        if (session == null) {
            logger.severe("Failed to create upload session.");
            return null;
        }

        try (InputStream fileStream = Files.newInputStream(localFile)) {
            var uploadTask = new LargeFileUploadTask<DriveItem>(
                    graphClient.getRequestAdapter(),
                    session,
                    fileStream,
                    fileSize,
                    SLICE_SIZE,
                    DriveItem::createFromDiscriminatorValue);

            IProgressCallback progress = (uploaded, total) -> {
                double percent = (double) uploaded / fileSize * 100;
                System.out.printf("\rProgress: %.1f%%   ", percent);
            };

            var result = uploadTask.upload(SLICE_SIZE, progress);
            System.out.println();

            if (result.isUploadSuccessful()) {
                return result.itemResponse;
            }
            logger.severe("Upload failed.");
            return null;
        }
    }
}
