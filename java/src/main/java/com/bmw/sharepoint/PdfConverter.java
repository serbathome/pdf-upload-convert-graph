package com.bmw.sharepoint;

import com.microsoft.graph.serviceclient.GraphServiceClient;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Set;
import java.util.logging.Logger;

public class PdfConverter {

    private static final Logger logger = Logger.getLogger(PdfConverter.class.getName());

    private static final Set<String> CONVERTIBLE_EXTENSIONS = Set.of(
            ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt");

    private final GraphServiceClient graphClient;
    private final String driveId;

    public PdfConverter(GraphServiceClient graphClient, String driveId) {
        this.graphClient = graphClient;
        this.driveId = driveId;
    }

    public boolean isConvertible(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        if (lastDot < 0) return false;
        return CONVERTIBLE_EXTENSIONS.contains(fileName.substring(lastDot).toLowerCase());
    }

    public SharePointUploader.UploadResult convertAndUpload(String sourceItemId, String pdfRemotePath) throws IOException {
        logger.info("Converting to PDF...");

        try (InputStream pdfStream = graphClient.drives().byDriveId(driveId)
                .items().byDriveItemId(sourceItemId)
                .content()
                .get(requestConfig -> requestConfig.queryParameters.format = "pdf")) {

            if (pdfStream == null) {
                logger.severe("PDF conversion returned no data.");
                return null;
            }

            byte[] pdfBytes = pdfStream.readAllBytes();
            String pdfItemId = "root:/" + pdfRemotePath + ":";

            try (var uploadStream = new ByteArrayInputStream(pdfBytes)) {
                var pdfItem = graphClient.drives().byDriveId(driveId)
                        .items().byDriveItemId(pdfItemId)
                        .content()
                        .put(uploadStream);

                if (pdfItem != null) {
                    logger.info("PDF conversion complete. Item ID: " + pdfItem.getId() + ", url: " + pdfItem.getWebUrl());
                    return new SharePointUploader.UploadResult(pdfItem.getId(), pdfItem.getWebUrl());
                }
                return null;
            }
        }
    }
}
