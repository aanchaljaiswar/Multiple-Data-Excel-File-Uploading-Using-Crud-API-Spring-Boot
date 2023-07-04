package com.springboot.excel.Service;

import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.springboot.excel.Entity.Excel;
import com.springboot.excel.Helper.ExcelHelper;
import com.springboot.excel.Repository.ExcelRepository;

@Service
public class ExcelService {

	@Autowired
	private ExcelRepository excelRepository;

	public void save(MultipartFile file) {
		try {
			List<Excel> excelList = ExcelHelper.excelToXCL(file.getInputStream());
			excelRepository.saveAll(excelList);
		} catch (IOException e) {
			throw new RuntimeException("Fail to store Excel data: " + e.getMessage());
		}
	}

	public List<Excel> getAllExcel() {
		return excelRepository.findAll();
	}
}
