package cn.tomleung.docx;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class test {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		InputStream is = new FileInputStream("C:\\DOC\\DOCX\\生产动态月报第六十五期.docx");
		XWPFDocument doc = new XWPFDocument(is);
//		Iterator<XWPFTable> it = doc.getTablesIterator();
//		while (it.hasNext()) {
//			XWPFTable table = it.next();
//			System.out.println(table.getRow(0).getCell(2).getText());
//		}
		XWPFTable table=doc.getTableArray(7);
		System.out.println(table.getRow(1).getCell(2).getText());
	}

}
