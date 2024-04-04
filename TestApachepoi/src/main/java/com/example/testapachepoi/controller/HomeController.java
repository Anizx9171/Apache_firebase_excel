package com.example.testapachepoi.controller;

import com.example.testapachepoi.firebase.FirebaseStorageService;
import com.example.testapachepoi.service.ImportExcelService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
@Controller
public class HomeController {
    @Autowired
    ImportExcelService importExcelService;
    @Autowired
    private FirebaseStorageService firebaseStorageService;
    @GetMapping("/")
    public String goIndex() {
        return "index";
    }
    @PostMapping("/upload")
    public String uploadExcel(@RequestParam("excelFile") MultipartFile mFile) {
        Path excelPathFile = null;
        try {
            String fileId = firebaseStorageService.uploadExcelFile(mFile);
            excelPathFile = firebaseStorageService.downloadExcelFile(fileId);
            if (importExcelService.readExcel(String.valueOf(excelPathFile))){
                System.out.println("ok");
            }else {
                System.out.println("er");
            }
            excelPathFile.toFile().delete();
        } catch (IOException e) {
            excelPathFile.toFile().delete();
            throw new RuntimeException(e);
        }
        return "redirect:/";
    }

    @GetMapping("/export")
    public void exportExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=bookDatabase.xls";
        response.setHeader(headerKey, headerValue);
        importExcelService.generateProductStatisticalExcel(response);
    }
}
