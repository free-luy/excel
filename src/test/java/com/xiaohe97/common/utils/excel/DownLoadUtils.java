package com.xiaohe97.common.utils.excel;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.support.rowset.SqlRowSet;

import com.xiaohe97.common.utils.SpringContextHolder;

public class DownLoadUtils {
	/**
	 * 下载Excel文件方法
	 * 
	 * @param response
	 * @param destFilePath
	 *            文件路径
	 * @param fileName
	 *            文件名称
	 * @throws IOException
	 */
	public static void download(HttpServletResponse response,
			String destFilePath, String fileName) throws IOException {
		OutputStream os = response.getOutputStream();
		try {
			File file = new File(destFilePath);
			// 以流的形式下载文件。
			InputStream fis = new BufferedInputStream(new FileInputStream(
					destFilePath));
			byte[] buffer = new byte[fis.available()];
			fis.read(buffer);
			fis.close();
			response.reset();
			// 设置response的Header
			/*
			 * response.addHeader("Content-Disposition", "attachment;filename="
			 * + new String(fileName.getBytes(), "UTF-8"));
			 */
			String filename = new String(fileName.getBytes("gbk"), "iso8859-1");
			response.setHeader("Content-Disposition", "attachment; filename=\""
					+ filename + "\"");
			response.addHeader("Content-Length", "" + file.length());
			response.setContentType("application/octet-stream");
			os.write(buffer);
			os.flush();
		} finally {
			if (os != null) {
				os.close();
			}
		}
	}

	public void print(String businessId, String businessTableName) {
		JdbcTemplate jdbcTemplate = SpringContextHolder.getBean("jdbcTemplate");
		String sql = "select id,table_name from form_table_basic_info where main_table_id = (select id from form_table_basic_info where instr ('"
				+ businessTableName + "',table_name) >0 ) ";
		SqlRowSet rowSet = jdbcTemplate.queryForRowSet(sql);
		String subTableId = null, tableName = null;
		if (rowSet.next()) {
			subTableId = rowSet.getString("id");
			tableName = rowSet.getString("table_name");
		}
		if (StringUtils.isBlank(subTableId) && StringUtils.isBlank(tableName)) {
			sql = "select * from business" + tableName + " b  inner join "
					+ businessTableName + " z on b.main_id = z.id";
			rowSet = jdbcTemplate.queryForRowSet(sql);
		}
	}
}
