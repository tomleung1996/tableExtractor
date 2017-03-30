package cn.tomleung.docx;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class ExtractLaw {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		process("c:/doc/docx", "");
	}

	public static void process(String folderPath, String resultPath) {
		File folder = new File(folderPath);
		File[] files = folder.listFiles();
		for (int i = 0; i < files.length; i++) {
			File file = files[i];
			try {
				extractLaw(file, resultPath);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	public static void extractLaw(File file, String resultPath) throws Exception {
		InputStream is = new FileInputStream(file);
		XWPFDocument doc = new XWPFDocument(is);
		Iterator<XWPFParagraph> it = doc.getParagraphsIterator();
		String year = "";
		Iterator<XWPFParagraph> it2 = doc.getParagraphsIterator();
		while (it2.hasNext()) {
			XWPFParagraph para = it2.next();
			Pattern p = Pattern.compile("\\（\\d{4}\\.\\d{1,2}\\）");
			Matcher m = p.matcher(para.getText());
			if (m.find()) {
				year = m.group().substring(1, 5) + "年";
				break;
			}
		}
		while (it.hasNext()) {
			String law;
			XWPFParagraph para = it.next();
			Pattern p = Pattern.compile("\\d+月\\d+日，[\\u4e00-\\u9fa5]+发[\\u4e00-\\u9fa5]+《\\W+》");
			Matcher m = p.matcher(para.getText());
			if (m.find()) {
				law = year + m.group();
				System.out.println(law);
			}

		}
		doc.close();
	}

}
