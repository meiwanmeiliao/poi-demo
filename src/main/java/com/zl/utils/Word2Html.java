package com.zl.utils;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;


/**
 * @author zhuolin
 * @program: poi-demo
 * @date 2018/11/27
 * @description: ${description}
 **/
public class Word2Html {
    /**
     * doc转html
     *
     * @param docPath           doc文件
     * @param htmlDirectoryPath 包存html文件路径
     * @return html文件
     * @throws Exception
     */
    private static String doc2Html (String docPath, String htmlDirectoryPath) throws Exception {
        String imagePathStr = htmlDirectoryPath + File.separator + "image" + File.separator;
        String targetFileName = htmlDirectoryPath + new File(docPath).getName() + ".html";
        File file = new File(imagePathStr);
        if (!file.exists()) {
            file.mkdirs();
        }
        HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(docPath));
        org.w3c.dom.Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
        //保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager((content, pictureType, name, width, height) -> {
            try (FileOutputStream out = new FileOutputStream(imagePathStr + name)) {
                out.write(content);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return "image/" + name;
        });
        wordToHtmlConverter.processDocument(wordDocument);
        org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(new File(targetFileName));
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        return targetFileName;
    }

    /**
     * docx转html
     *
     * @param wordPath
     * @param htmlDirectoryPath
     * @return
     */
    private static String docx2Html (String wordPath, String htmlDirectoryPath) {
        // word文档路径为空
        if (wordPath == null) {
            System.err.println("error:path of word document is null");
            return null;
        }
        File wordFile = new File(wordPath).getAbsoluteFile();
        File htmlFile = new File(htmlDirectoryPath + File.separator + wordFile.getName() + ".html");
        try {
            // 输入流
            InputStream inputStream = new FileInputStream(wordFile);
            // 读取word文档
            XWPFDocument document = new XWPFDocument(inputStream);
            // 关闭输入流
            inputStream.close();
            // 创建选项
            XHTMLOptions options = XHTMLOptions.create();
            // 设置图片文件夹保存的路径以及文件夹名称
            options.setExtractor(new FileImageExtractor(htmlFile.getParentFile()));
            // 输出流
            OutputStream outputStream = new FileOutputStream(htmlFile);
            // word文档转html
            XHTMLConverter.getInstance().convert(document, outputStream, options);
            // 关闭输出流
            outputStream.close();
            // 关闭文档
            document.close();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
        return htmlFile.getAbsolutePath();
    }

    public static void main (String argv[]) {
        try {
            System.out.println(doc2Html("C:\\Users\\zhuolin\\Desktop\\aaa.doc", "D:\\html\\"));
            System.out.println(docx2Html("C:\\Users\\zhuolin\\Desktop\\aaa.docx", "D:\\html\\"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
