package file.upload;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class TransFile {
	
	
	public void transFile(String str,String fileName){
		
		File outFile=new File(str+"/"+fileName);//合成后的文件
		if(outFile.exists()){
			outFile.delete();
		}
		
		File root=new File(str);
		int a=root.listFiles().length;
		
		for(int i=0;i<a;i++){
			InputStream in=null;
			OutputStream out=null;
			try {
				in=new FileInputStream(str+"/"+i);
				out=new FileOutputStream(str+"/"+fileName, true);
				int b=0;
				while((b=in.read())!=-1){
					out.write(b);
				}
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}finally{
				try {
					out.close();
					in.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		System.out.println("文件合并成功");
	}
	
     public static void main(String[] args) {
		TransFile tr=new TransFile();
		tr.transFile("D:/temporaryFiles/089cd103d8566e83957a20ff7b294645", "sogou_pinyin_50f.exe");
	}
}
