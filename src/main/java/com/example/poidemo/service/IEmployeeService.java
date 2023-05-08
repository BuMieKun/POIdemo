package com.example.poidemo.service;

import org.springframework.web.multipart.MultipartFile;

public interface IEmployeeService {
    String uploadExcel(MultipartFile file);
}
