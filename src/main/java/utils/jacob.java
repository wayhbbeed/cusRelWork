package main.java;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
//import test.wordBean;
public class Jacob {
	 // word文档
	 private static Dispatch doc;
	 // word运行程序对象
	 private ActiveXComponent word;
	 // 所有word文档集合
	 private Dispatch documents;
	 // 选定的范围或插入点
	 private Dispatch selection;
	 private boolean saveOnExit = true;
	 
	 private ArrayList alist=null;
	 
	 public Jacob() throws Exception {
		  if (word == null) {
			  word = new ActiveXComponent("Word.Application");
			  word.setProperty("Visible", new Variant(false)); // 不可见打开word
			  word.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
		  }
		  if (documents == null)
			  documents = word.getProperty("Documents").toDispatch();
		 }

//创建一个新的word文档
	 public void createNewDocument() {
		  doc = Dispatch.call(documents, "Add").toDispatch();
		  selection = Dispatch.get(word, "Selection").toDispatch();
		 }
 
//打开一个已存在的文档
	 public void openDocument(String docPath) {
		  createNewDocument();
		  doc = Dispatch.call(documents, "Open", docPath).toDispatch();
		  selection = Dispatch.get(word, "Selection").toDispatch();
	 }

 //获得指定的单元格里数据
	 public String getTxtFromCell(int tableIndex, int cellRowIdx, int cellColIdx) {
	  // 所有表格
		  Dispatch tables = Dispatch.get(doc, "Tables").toDispatch(); 
		  // 要填充的表格
		  Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		  Dispatch rows = Dispatch.call(table, "Rows").toDispatch();
		  Dispatch columns = Dispatch.call(table, "Columns").toDispatch();
		  Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),new Variant(cellColIdx)).toDispatch();
		  Dispatch Range=Dispatch.get(cell,"Range").toDispatch();
		//  System.out.println(Dispatch.get(Range,"Text").toString());
		  Dispatch.call(cell, "Select");
		  String ret = "";
		  ret = Dispatch.get(selection, "Text").toString();
		  ret = ret.substring(0, ret.length() - 2); // 去掉最后的回车符;
		  return ret;
	 }
 //关闭
	 public void closeDocumentWithoutSave() {
		  if (doc != null) {
		   Dispatch.call(doc, "Close", new Variant(false));
		   doc = null;
		  }
	 }
 
 //关闭全部应用
	 public void close() {
		 //closeDocument();
		 if (word != null) {
			 Dispatch.call(word, "Quit");
			 word = null;
		 }
			 selection = null;
			 documents = null;
	 }
	 
	 //遍历文件夹下的所有文件路径
	 public  ArrayList getFilePath(File filedir) throws Exception{
		  File[] file = filedir.listFiles();
		  alist=new ArrayList();
		  for(int i=0; i<file.length; i++){
          //获取绝对路径
		   String filepath=file[i].getAbsolutePath();
//		   System.out.println(file[i].getAbsolutePath());
		   alist.add(filepath);
//		   System.out.println(alist.get(i));
		   if(file[i].isDirectory()){
		    try{
		    	getFilePath(file[i]);
		    }catch(Exception e){}
		   }
		  }
		  return alist;
		 }
	 
	 //遍历文件夹下的所有文件名
	 public  ArrayList getFileName(File filedir) throws Exception{
		  alist=new ArrayList();
		  File[] file = filedir.listFiles();
		  for(int i=0; i<file.length; i++){
		   //获取文件名
		   String filename=file[i].getName().substring(0, file[i].getName().indexOf("."));
		   alist.add(filename);
//		   System.out.println(alist.get(i));
		   if(file[i].isDirectory()){
		    try{
		    	getFileName(file[i]);
		    }catch(Exception e){}
		   }
		  }
		  return alist;
		 }
	 
//测试方法
 public static void main(String[] args)throws Exception{
	  Jacob word = new Jacob(); 
	  
	  File myfiledir = new File("C:\\list");
      String filepaths=word.getFilePath(myfiledir).get(1).toString();
      String filenames=word.getFileName(myfiledir).get(1).toString();
	  System.out.println(filepaths+filenames);
      
	  // 打开word
	  word.openDocument(filepaths);
	  
	  ArrayList warningList = new ArrayList();
	  
	  // 所有表格
	  Dispatch tables = Dispatch.get(doc, "Tables").toDispatch(); 
	  // 获取表格数目
	  int tablesCount = Dispatch.get(tables,"Count").toInt();
	  System.out.println("tablesCount"+"  "+tablesCount);
	  // 循环获取表格
	  for(int i=1;i<=tablesCount;i++)
		  {
		   // 生产warningBean
		   // warningBean warning = new warningBean();
		   // 获取第i个表格
		   Dispatch table = Dispatch.call(tables, "Item", new Variant(i)).toDispatch();
		   // 获取该表格所有行
		   Dispatch rows = Dispatch.call(table, "Rows").toDispatch();
		   // 获取该表格所有列
		   Dispatch columns = Dispatch.call(table, "Columns").toDispatch();
		   // 获取该表格行数
		   int rowsCount = Dispatch.get(rows,"Count").toInt();
		   System.out.println("rowsCount"+"  "+rowsCount);
		   // 获取该表格列数
		   int columnsCount = Dispatch.get(columns,"Count").toInt();
		   System.out.println("columnsCount"+"  "+columnsCount);
		   
			   // 循环遍历行
			   for(int j=1;j<=rowsCount;j++)
			   {
			    // 获取第i个表格第j行第一列标题
//			    String title = word.getTxtFromCell(i, j, 1);
			    String titles = word.getTxtFromCell(1, j, 2);
				System.out.println(titles.toString().trim());
			   }
				  }  
				  // 关闭该文档
				  word.closeDocumentWithoutSave();
				  // 关闭word
				  word.close();
				  } 
}


