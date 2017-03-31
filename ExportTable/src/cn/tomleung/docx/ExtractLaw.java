package cn.tomleung.docx;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class ExtractLaw {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		ArrayList<String> date=new ArrayList<String>();
		ArrayList<String> who=new ArrayList<String>();
		ArrayList<String> laws=new ArrayList<String>();
		ArrayList<ArrayList<String>> list = new ArrayList<ArrayList<String>>();
		list.add(date);
		list.add(who);
		list.add(laws);
		try {
			process("D:\\广东建工集团-管理+互联网项目\\动态月报\\转换成DOCX并修复第一第二期问题的月报", "D:\\广东建工集团-管理+互联网项目\\动态月报\\法规抽取",list);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void process(String folderPath, String resultPath,ArrayList<ArrayList<String>> list) throws Exception {
		File folder = new File(folderPath);
		File[] files = folder.listFiles();
		for (int i = 0; i < files.length; i++) {
			File file = files[i];
			try {
				extractLaw(file, resultPath,list);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		StringBuilder sb=buildString(list);
		FileWriter fw = new FileWriter(
				resultPath + "\\\\Law.csv");
		BufferedWriter bw = new BufferedWriter(fw);
		bw.write(sb.toString());
		bw.close();

	}

	public static void extractLaw(File file, String resultPath,ArrayList<ArrayList<String>> list) throws Exception {
		if (file.isDirectory()||!file.getName().substring(file.getName().lastIndexOf("."), file.getName().length()).equals(".docx")
				&& !file.getName().substring(file.getName().lastIndexOf("."), file.getName().length()).equals(".DOCX")
				){
			return;
		}
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
			Pattern p = Pattern.compile("\\d+月\\d+日，[\\u4e00-\\u9fa5]+发[\\u4e00-\\u9fa5]+《.+》");
			Matcher m = p.matcher(para.getText());
			if (m.find()) {
				law = year + m.group();
				int pos1=law.indexOf("，");
				p=Pattern.compile("部|委|协会|局|厅");
				m=p.matcher(law);
				if(!m.find())
					continue;
				int pos2=m.start();
				int pos3=law.indexOf("《");
				//0,pos1
				//pos1+1,pos2+1
				//pos3
				list.get(0).add(law.substring(0,pos1));
				list.get(1).add(law.substring(pos1+1,pos2+1));
				list.get(2).add(law.substring(pos3));
			}
		}
		doc.close();
	}
	
	public static StringBuilder buildString(ArrayList<ArrayList<String>> list){
		StringBuilder sb=new StringBuilder("");
		int size=list.get(0).size();
		for(int i=0;i<size;i++){
			sb.append("\"");
			sb.append(list.get(0).get(i));
			sb.append("\",\"");
			sb.append(list.get(1).get(i));
			sb.append("\",\"");
			sb.append(list.get(2).get(i));
			sb.append("\"\r\n");
		}
		return sb;
	}

}
