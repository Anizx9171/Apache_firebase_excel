package com.example.testapachepoi.firebase;
import com.google.auth.oauth2.GoogleCredentials;
import com.google.cloud.storage.Storage;
import com.google.cloud.storage.StorageOptions;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.ClassPathResource;
import java.io.IOException;

@Configuration
public class FirebaseConfig {

    @Bean
    public Storage storage() throws IOException {
        GoogleCredentials credentials = GoogleCredentials.fromStream(new ClassPathResource("ojt2-d1c51-firebase-adminsdk-lduy5-8d89e7aba7.json").getInputStream());
        Storage storage = StorageOptions.newBuilder().setCredentials(credentials).build().getService();
        return storage;
    }
}
