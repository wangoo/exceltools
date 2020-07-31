package org.wang.excelUtils.service;

import com.alibaba.excel.EasyExcel;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
    public static void unionWorkBook(String path,String fileType){
        File folder = new File(path);
        if(!folder.isDirectory()){
            return;
        }
        List<Map<String,String>> dataMap = new ArrayList<>();
        String[] fileNames = folder.list();
        List<List<String>> header = null;
        List<List<String>> body = new ArrayList<>();
        boolean   hasHeader = false;

        for(String filename : fileNames){

            if(!filename.endsWith(fileType)){
                continue;
            }

            filename= Paths.get(path,filename).toString();
            DataListener dataListener = new DataListener() ;
            List<Map<String,String>> results =  EasyExcel.read(filename,dataListener).sheet().doReadSync();
             for(Map<String,String> map:results){
                 List<String> row = new ArrayList<>();
                 for (int i = 0; i <map.size() ; i++) {
                     row.add(map.get(i));
                 }
                 body.add(row);
             }
             if(!hasHeader){
                 header =dataListener.header;
                 hasHeader =true;
             }
        }
        System.out.println(header);
        System.out.println(body);
        EasyExcel.write( Paths.get(path,"合并文件"+fileType).toString()).head(header).sheet("模板").doWrite(body);
    }
}
