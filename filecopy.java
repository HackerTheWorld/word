package webproject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.Iterator;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.XWPFConverterException;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.*;




public class filecopy {

	public static void main(String[] args){
//	  filecheng("D:/hc20170328x 出口货物明细.doc",16,3);
	   try {
		   filechengtype("在业务操作模块上.docx","hc20170328x 出口货物明细.html","D://");
	    } catch (Exception e) {
		// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void filecheng(String str,int row,int cell){
		
		File file = new File(str);
		
		if(file.getName().endsWith(".doc")){
			doc(file,row,cell);
		}else if(file.getName().endsWith(".docx")){
			docx(file,row,cell);
		}
		
	}
	
	public static void doc(File file,int row,int cell){
		  
		  try{
			  FileInputStream fis = new FileInputStream(file);
			  HWPFDocument doc = new HWPFDocument(fis);
			  Range range =doc.getRange();  
	          TableIterator table=new TableIterator(range);
	          while(table.hasNext()){
	        	  Table ta=table.next(); 
	        	  TableRow tr= ta.getRow(row);
	        	  TableCell td=tr.getCell(cell);
	        	  for(int k=0;k<td.numParagraphs();k++){  
	        		  org.apache.poi.hwpf.usermodel.Paragraph para=td.getParagraph(k);
	        		  System.out.println(para.text().trim());
	        	  }
	          }
		  }catch(Exception e){
			  e.printStackTrace();
		  }
		 
		
          
	}
	
	public static void docx(File file,int row,int cell){
		  try{
			  FileInputStream fis = new FileInputStream(file);
			  XWPFDocument docx = new XWPFDocument(fis);
			  Iterator<XWPFTable> table =docx.getTablesIterator();  
	          while(table.hasNext()){
	        	  XWPFTable ta=table.next(); 
	        	  XWPFTableRow tr= ta.getRow(row);
	        	  XWPFTableCell td=tr.getCell(cell);
	        	  String s1=td.getText();
	        	  System.out.println(s1);
	          }
		  }catch(Exception e){
			  e.printStackTrace();
		  }
	}
	
	public static void filechengtype(String fileName,String htmlName,String filepath) throws IOException, TransformerException, ParserConfigurationException{
		File file = new File(filepath+fileName);
		WordToHtmlAction doc=new WordToHtmlAction();
		System.out.println(file.getName());
		if(file.getName().endsWith(".doc")){
			  doc.Word2003ToHtml(fileName, htmlName, filepath);
		}else if(file.getName().endsWith(".docx")){
			  doc.Word2007ToHtml(fileName, htmlName, filepath);
		}
	}
	
	
	
	
}
