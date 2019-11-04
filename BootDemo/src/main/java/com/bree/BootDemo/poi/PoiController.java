package com.bree.BootDemo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.bree.BootDemo.common.ExcelFile;

@RestController
@RequestMapping("/poi")
public class PoiController {

	@RequestMapping("/download")
	public void downloadFile(HttpServletResponse response) throws IOException {
		try {
			ClassPathResource resource = new ClassPathResource("poi.xlsx");
			File file = resource.getFile();
			String filename = resource.getFilename();
			InputStream inputStream = new FileInputStream(file);
			// 强制下载不打开
			response.setContentType("application/force-download");
			OutputStream out = response.getOutputStream();
			// 使用URLEncoder来防止文件名乱码或者读取错误
			response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(filename, "UTF-8"));
			int b = 0;
			byte[] buffer = new byte[1000000];
			while (b != -1) {
				b = inputStream.read(buffer);
				if (b != -1)
					out.write(buffer, 0, b);
			}
			inputStream.close();
			out.close();
			out.flush();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * excel生成下载
	 * 
	 * @param response
	 * @return
	 * @throws Exception
	 */
	@GetMapping(value = "/createExcel")
	public String createExcel(HttpServletResponse response) throws Exception {
		Map<String, Object> excelMap = new HashMap<>();
		// 1.设置Excel表头
		List<String> headerList = new ArrayList<>();
		headerList.add("用户id");
		headerList.add("用户名");
		headerList.add("性别");
		headerList.add("身份证号");
		headerList.add("注册时间");
		excelMap.put("header", headerList);

		// 2.是否需要生成序号，序号从1开始(true-生成序号 false-不生成序)
		boolean isSerial = false;
		excelMap.put("isSerial", isSerial);

		// 3.sheet名
		String sheetName = "统计表";
		excelMap.put("sheetName", sheetName);

		// 4.需要放入Excel中的数据
		Map<String, Object> map = new HashMap<>();
		map.put("gender", "男");
		// List<User> rows = userMapper.selectUserInfo(map);

		List<List<String>> data = new ArrayList<>();
		// List<User> rows = userMapper.selectUserInfo(map);
		// SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		// 此处往cell中set的值写死了，可以取数据库中值进行赋值
		for (int i = 0; i < 1; i++) {
			// 所有的数据顺序必须和表头一一对应
			// list存放每一行的数据（让所有的数据类型都转换成String，这样就无需担心Excel表格中数据不对）
			List<String> list = new ArrayList<>();
			list.add(String.valueOf(12345l));
			list.add(String.valueOf("1123"));
			list.add(String.valueOf("男"));
			list.add(String.valueOf("china"));
			list.add(String.valueOf("2019-01-01"));
			// data存放全部的数据
			data.add(list);
		}
		excelMap.put("data", data);

		// Excel文件内容设置
		HSSFWorkbook workbook = ExcelFile.createExcel(excelMap);

		String fileName = "导出excel例子.xls";

		// 生成excel文件
		ExcelFile.buildExcelFile(fileName, workbook);
		// 浏览器下载excel
		ExcelFile.buildExcelDocument(fileName, workbook, response);

		return "success";

	}
}
