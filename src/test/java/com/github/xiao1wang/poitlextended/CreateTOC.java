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
public class CreateTOC {

	public static void main(String[] args) throws Exception {
		FileInputStream fileInputStream = new FileInputStream(ChartTest.class.getClassLoader().getResource("templates/test_doc.docx").getPath());
		XWPFDocument doc = new XWPFDocument(fileInputStream);
		CustomerTOC.handGenerateTOC(doc, "toc");
		OutputStream out = new FileOutputStream("d:\\my_doc.docx");
		doc.write(out);
		out.close();
	}
}
