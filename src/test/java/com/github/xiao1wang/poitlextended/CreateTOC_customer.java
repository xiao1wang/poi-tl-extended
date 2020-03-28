package com.github.xiao1wang.poitlextended;

import com.github.xiao1wang.poitlextended.util.CustomerTOC;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * @TODO 写点注释
 * @Author : wangyahui
 * @Date: 2020-03-27 09:44
 */
public class CreateTOC_customer {

	public static void main(String[] args) throws Exception {
		FileInputStream fileInputStream = new FileInputStream(ChartTest.class.getClassLoader().getResource("templates/test_toc.docx").getPath());
		XWPFDocument doc = new XWPFDocument(fileInputStream);
		int maxLevel = 3;
		CustomerTOC.automaticGenerateTOC(maxLevel, "toc", doc);
		OutputStream out = new FileOutputStream("d:\\my_doc_customer.docx");
		doc.write(out);
		out.close();
	}

}
