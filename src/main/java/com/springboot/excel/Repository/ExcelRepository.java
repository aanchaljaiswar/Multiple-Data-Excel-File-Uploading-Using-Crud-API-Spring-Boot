package com.springboot.excel.Repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.springboot.excel.Entity.Excel;

public interface ExcelRepository extends JpaRepository<Excel, Long> {

}
