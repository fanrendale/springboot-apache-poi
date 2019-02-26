package com.xjf.apachepoitest.controller;

import io.swagger.annotations.ApiOperation;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 使用poi解析Word文档，格式为doc或者docx
 * @author xjf
 * @date 2019/2/25 16:28
 */
@RestController
public class PoiWordController {

    @ApiOperation(value = "解析word文档",notes = "按段落解析word文档")
    @RequestMapping(value = "upload",method = RequestMethod.POST)
    public Map uploadFile(@RequestParam(value = "file",required = true)MultipartFile file){
        //获取文件名
        String textFileName = file.getOriginalFilename();
        Map wordMap = new LinkedHashMap();      //创建一个map对象存放word中的内容

        try {
            if (textFileName.endsWith(".doc")){     //判断文件格式
                InputStream fis = file.getInputStream();
                //使用HWPF组件中的WordExtractor类从Word文档中提取文本或段落
                WordExtractor wordExtractor = new WordExtractor(fis);
                int i = 1;
                for(String words : wordExtractor.getParagraphText()){//获取段落内容
                    System.out.println(words);
                    wordMap.put("DOC文档，第（"+i+"）段内容",words);
                    i++;
                }
                fis.close();
            }

            if(textFileName.endsWith(".docx")){
                File uFile = new File("tempFile.docx");//创建一个临时文件
                if(!uFile.exists()){
                    uFile.createNewFile();
                }
                FileCopyUtils.copy(file.getBytes(), uFile);//复制文件内容
                //包含所有POI OOXML文档类的通用功能，打开一个文件包。
                OPCPackage opcPackage = POIXMLDocument.openPackage("tempFile.docx");
                //使用XWPF组件XWPFDocument类获取文档内容
                XWPFDocument document = new XWPFDocument(opcPackage);
                List<XWPFParagraph> paras = document.getParagraphs();
                int i=1;
                for(XWPFParagraph paragraph : paras){
                    String words = paragraph.getText();
                    System.out.println(words);
                    wordMap.put("DOCX文档，第（"+i+"）段内容",words);
                    i++;
                }
                uFile.delete();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(wordMap);
        return wordMap;
    }
}
