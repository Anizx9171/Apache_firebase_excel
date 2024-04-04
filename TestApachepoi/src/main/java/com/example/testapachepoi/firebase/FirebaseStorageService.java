package com.example.testapachepoi.firebase;
import com.google.cloud.storage.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@Service
public class FirebaseStorageService {
    @Autowired
    private Storage storage;
    private final String bucketName = "ojt2-d1c51.appspot.com";
    public String uploadExcelFile(MultipartFile file) throws IOException {

        String fileName = file.getOriginalFilename();

        BlobId blobId = BlobId.of(bucketName, fileName);
        BlobInfo blobInfo = BlobInfo.newBuilder(blobId).build();
        Blob blob = storage.create(blobInfo, file.getBytes());

        return blob.getBlobId().getName();
    }

    public Path downloadExcelFile(String fileId) throws IOException {
        // Get bucket
        Bucket bucket = storage.get(bucketName);
        // Get path to resources/upload directory
        String uploadDirectoryPath = new ClassPathResource("upload").getFile().getAbsolutePath();
        // Generate file path
        Path filePath = Paths.get(uploadDirectoryPath, fileId);
        // Download data to file
        bucket.get(fileId).downloadTo(filePath);
        return filePath;
    }
}
