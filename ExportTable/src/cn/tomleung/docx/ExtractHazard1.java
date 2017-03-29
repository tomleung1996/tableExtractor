package cn.tomleung.docx;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ExtractHazard1 {

	public static void main(String[] args) {
		Scanner in = new Scanner(System.in);
		System.out.println("此程序可以将DOCX格式中所想要的表格导出为CSV格式");
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
			File file=files[i];
			try {
				extractToCSVFromDocx(file, resultPath);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	public static void extractToCSVFromDocx(File file, String resultPath) throws Exception {
		if (file.isDirectory()||!file.getName().substring(file.getName().lastIndexOf("."), file.getName().length()).equals(".docx")
				&& !file.getName().substring(file.getName().lastIndexOf("."), file.getName().length()).equals(".DOCX")
				){
			return;
		}
		InputStream is = new FileInputStream(file);
		XWPFDocument doc = new XWPFDocument(is);
		Iterator<XWPFTable> it = doc.getTablesIterator();
		StringBuilder str = new StringBuilder("");
		String title = "";
		Iterator<XWPFParagraph> it2 = doc.getParagraphsIterator();
		while (it2.hasNext()) {
			XWPFParagraph para = it2.next();
			Pattern p = Pattern.compile("\\（\\d{4}\\.\\d{1,2}\\）");
			Matcher m = p.matcher(para.getText());
			if (m.find()) {
				title = m.group();
				title = title.replaceAll("（", "");
				title = title.replaceAll("）", "");
				break;
			}
		}
		while (it.hasNext()) {
			XWPFTable table = it.next();
			List<XWPFTableRow> rows = table.getRows();
			if (rows.get(0).getCell(2) == null || !rows.get(0).getCell(2).getText().contains("深基坑")) {
				continue;
			}
			int count = rows.size();
			for (int i = 2; i < count; i++) {
				List<XWPFTableCell> cells = rows.get(i).getTableCells();
				if (i > 0 && cells.get(0).getText().equals("序号"))
					continue;
				for (XWPFTableCell c : cells) {
					str.append("\"" + c.getText() + "\",");
				}
				if(i==0){
					str.append("\"日期\",");
				}else if(i!=0){
					str.append("\""+title+"\",");
				}
				str.deleteCharAt(str.length() - 1);
				str.append("\r\n");
			}
		}
		if(str.length()<=1){
			doc.close();
			return;
		}
		title=title.replaceAll("\\.", "_");
		Pattern p = Pattern.compile("(\"\",){10,11}\"\"");
		Matcher m = p.matcher(str);
		if (m.find()){
			str = str.delete(m.start(), m.end()+1);
		}
//		FileWriter fw = new FileWriter(
//				resultPath + "\\\\" + title + "-" + file.getName().substring(0, file.getName().length() - 5) + ".csv");
		FileWriter fw = new FileWriter(
				resultPath + "\\\\Project_Hazard1_" + title + ".csv");
		BufferedWriter bw = new BufferedWriter(fw);
		bw.write(str.toString());
		bw.close();
		doc.close();
		str = new StringBuilder("");
		doc.close();
	}

}
