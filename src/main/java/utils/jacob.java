package main.java.utils;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import main.java.Cus;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
//import test.wordBean;
public class Words {
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
	 private ArrayList<Cus> al1=null;
	 private Cus cus=null;
	 
	 public Words() throws Exception {
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
	 //获取文件夹下文件数量
	 public  int getFileCount(File filedir) throws Exception{
		  File[] file = filedir.listFiles();
		  int filecount=file.length;
		  return filecount;
		 }
	 
	 //获取表格中的所有有效内容
	 public ArrayList<Cus> getFileContent() throws Exception{
		 
		  File myfiledir = new File("C:\\list");
		  File[] file = myfiledir.listFiles();
		  
		  al1=new ArrayList<Cus>();
		  cus=new Cus();
		  
		  for (int i = 0; i < file.length; i++) {
			  System.out.println(file.length);
			  if (!file[i].isHidden()) {
				  
				  String filepaths=getFilePath(myfiledir).get(i).toString();
				  String filenames=getFileName(myfiledir).get(i).toString();
				  openDocument(filepaths);
				  // 所有表格
				  Dispatch tables = Dispatch.get(doc, "Tables").toDispatch(); 
				  // 获取第1个表格
				  Dispatch table = Dispatch.call(tables, "Item", new Variant(1)).toDispatch();
				  // 获取表格值（固定10行）
			      cus=new Cus();
			      cus.setId(filenames);
			      cus.setProvince(getTxtFromCell(1, 1, 2));
			      cus.setProvince_manager(getTxtFromCell(1, 2, 2));
			      cus.setCustom(getTxtFromCell(1, 3, 2));
			      cus.setBusiness(getTxtFromCell(1, 4, 2));
			      cus.setReceiver(getTxtFromCell(1, 5, 2));
			      cus.setBill_info(getTxtFromCell(1, 6, 2));
			      cus.setInvoice_info(getTxtFromCell(1, 7, 2));
			      cus.setCus_info(getTxtFromCell(1, 8, 2));
			      cus.setGoods(getTxtFromCell(1, 9, 2));
			      cus.setRequirement(getTxtFromCell(1, 10, 2));
			      closeDocumentWithoutSave();
			}else{
				  System.out.println("请关闭所有Word文档，并结束进程！");
				  break;
			}
			      al1.add(cus);
		    }
		          close();
		          return al1;
	 }
	 
//测试方法
// public static void main(String[] args)throws Exception{
//	  Jtest word = new Jtest(); 
//	  ArrayList<Cus> list=new ArrayList<Cus>();
//	  list=word.getFileContent();
//	  Iterator<Cus> it=list.iterator();
//	  while (it.hasNext()) {
//		System.out.println(it.next());
//		
//	}
//		} 
}


