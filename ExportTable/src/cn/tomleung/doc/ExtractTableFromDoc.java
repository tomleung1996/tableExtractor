package cn.tomleung.doc;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExtractTableFromDoc {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Scanner in = new Scanner(System.in);
		System.out.println("请输入要读取的文件夹路径（请确保有访问权限）：");
		String from = in.nextLine();
		from = from.replaceAll("\\\\", "\\\\\\\\");
		System.out.println("请输入要输出的文件夹路径（请确保有访问权限）：");
		String to = in.nextLine();
		to = to.replaceAll("\\\\", "\\\\\\\\");
		process(from, to);
		System.out.println("导出完成");
		in.close();

	}

	public static void process(String folderPath, String resultPath) {
		File folder = new File(folderPath);
		File[] files = folder.listFiles();
		for (int i = 0; i < files.length; i++) {
			try {
				extractToCSVFromDoc(files[i], resultPath);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	public static void extractToCSVFromDoc(File file, String resultPath) throws Exception {
		if (file.isDirectory())
			return;
		String title = "";
		StringBuilder str = new StringBuilder("");
		String filename = file.getAbsolutePath();
		InputStream fis = new FileInputStream(filename);
		POIFSFileSystem fs = new POIFSFileSystem(fis);
		HWPFDocument doc = new HWPFDocument(fs);
		Range range = doc.getRange();
		boolean intable = false;
		boolean inrow = false;
		boolean reachStart = false;
		boolean titleSaved = false;
		// [\\u4e00-\\u9fa5]+
		for (int i = 0; i < range.numParagraphs(); i++) {
			Paragraph par = range.getParagraph(i);

			Pattern p = Pattern.compile("\\（\\d{4}\\.\\d{1,2}\\）");
			Matcher matcher = p.matcher(par.text());
			if (!matcher.find() && reachStart == false) {
				continue;// 还没到指定表格，继续
			}
			// 到了
			if (titleSaved == false) {
				title = matcher.group();// 保存下名字
				title = title.replaceAll("（", "");
				title = title.replaceAll("）", "");
				titleSaved = true;
			}
			reachStart = true;

			p = Pattern.compile(
					"\\u6709\\u8d28\\u91cf\\u5b89\\u5168\\u8bc4\\u4ef7\\u5206\\u6570\\u7684\\u9879\\u76ee\\u662f\\u7eb3\\u5165\\u5e02\\u5efa\\u59d4\\u73b0\\u573a\\u8bc4\\u4ef7\\u7684\\u5e7f\\u5dde\\u5e02\\u9879\\u76ee");
			matcher = p.matcher(par.text());
			if (matcher.find())
				break;// 匹配到表格下方的特定注释则表示表格读取完，跳出

			if (par.isInTable()) {
				if (!intable) {
					intable = true;
				}
				if (!inrow) {
					inrow = true;
				}
				if (par.isTableRowEnd()) {
					str.delete(str.length() - 1, str.length());
					str.append(par.text().substring(0, par.text().length() - 1) + "\r\n");
					inrow = false;
				} else {
					str.append(par.text().substring(0, par.text().length() - 1) + ",");
				}
			}
		}
		Pattern p = Pattern.compile("\\r\\n[,]{10,11}");
		Matcher matcher = p.matcher(str);
		if (matcher.find())
			str = new StringBuilder(matcher.replaceFirst(""));
		str.insert(0, title + "\r\n");
		str.delete(str.length() - 3, str.length());
		FileWriter fw = new FileWriter(resultPath + "\\\\" + title + "-"
				+ file.getName().substring(0, file.getName().length() - 4) + ".csv");
		BufferedWriter bw = new BufferedWriter(fw);
		bw.write(str.toString());
		bw.close();
		doc.close();
		str = new StringBuilder("");
		return;
	}

}
