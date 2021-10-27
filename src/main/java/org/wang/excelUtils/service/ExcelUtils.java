package org.wang.excelUtils.service;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
    public static void unionWorkBook(String path,String fileType,boolean unionAllSheet){
        File folder = new File(path);
        if(!folder.isDirectory()){
            return;
        }

        String[] fileNames = folder.list();
        List<List<String>> header = null;
        List<List<String>> body = new ArrayList<>();
        boolean   hasHeader = false;

        for(String filename : fileNames){

            if(!filename.endsWith(fileType)){
                continue;
            }

            String dataSourceFilename = filename;
            filename= Paths.get(path,filename).toString();
            DataListener dataListener = new DataListener() ;



            List<Map<String,String>> results = null;

            if(unionAllSheet){
                results =   EasyExcel.read(filename,dataListener).doReadAllSync();
            }else{
                results =EasyExcel.read(filename,dataListener).sheet().doReadSync();
            }

             for(Map<String,String> map:results){
                 List<String> row = new ArrayList<>();
                 for (int i = 0; i <map.size() ; i++) {
                     row.add(map.get(i));
                 }
                 row.add(dataSourceFilename);
                 body.add(row);
             }
             if(!hasHeader){
                 header =dataListener.header;
                 List<String> dataSourceHeader = new ArrayList<>();
                 dataSourceHeader.add("数据来源文件");
                 header.add(dataSourceHeader);
                 hasHeader =true;
             }
        }
        EasyExcel.write( Paths.get(path,"文件"+fileType).toString()).head(header).sheet("模板").doWrite(body);
    }
}
