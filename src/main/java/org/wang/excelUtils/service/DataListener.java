package org.wang.excelUtils.service;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class DataListener extends AnalysisEventListener<Map<String,String>> {

    private static final int BATCH_COUNT = 5;
    List<Map<String,String>> list = new ArrayList<Map<String,String>>();
    List<List<String>> header = new ArrayList<>();

    public DataListener() {
        super();
    }


    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        if(headMap==null||headMap.isEmpty()){
            return;
        }

        for (int i = 0; i < headMap.size(); i++) {
            ArrayList<String> headerRow = new ArrayList();
            headerRow.add(headMap.get(i));
            header.add(headerRow);
        }

    }

    @Override
    public void invoke(Map<String, String> stringStringMap, AnalysisContext analysisContext) {

    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println(list.toString());
    }
}

