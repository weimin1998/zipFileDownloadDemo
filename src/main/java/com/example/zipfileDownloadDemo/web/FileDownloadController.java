package com.example.zipfileDownloadDemo.web;

import com.example.zipfileDownloadDemo.TemplateUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Controller
public class FileDownloadController {

    @GetMapping("/zipFile")
    public void zipFile(HttpServletResponse response) throws IOException {
        String zipFileName = "压缩文件.zip";
        response.setHeader("Content-Disposition", "attachment;fileName=" + new String(zipFileName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1));

        ZipOutputStream zipOutputStream = new ZipOutputStream(response.getOutputStream());

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        // 准备数据
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("name", "魏敏");
        dataMap.put("fale", "男");
        dataMap.put("age", "22");

        String docPath = "doc/word文档.docx";

        // 加载文件
        ClassPathResource classPathResource = new ClassPathResource(docPath);
        XWPFDocument xwpfDocument = new XWPFDocument(classPathResource.getInputStream());
        TemplateUtils.replacePlaceholder(xwpfDocument, dataMap);

        Map<String, String> fileMap = getFileUrls();


        zipOutputStream.putNextEntry(new ZipEntry("填充模板" + File.separator + "word文档.docx"));
        xwpfDocument.write(byteArrayOutputStream);
        zipOutputStream.write(byteArrayOutputStream.toByteArray());


        for (Map.Entry<String, String> entry : fileMap.entrySet()) {
            UrlResource urlResource = new UrlResource(entry.getValue());
            try (InputStream inputStream = urlResource.getInputStream()) {
                zipOutputStream.putNextEntry(new ZipEntry("上传文件" + File.separator + entry.getKey()));
                IOUtils.copy(inputStream, zipOutputStream);
            } catch (IOException e) {
                //log.info("下载文件失败！");
            }
        }

        zipOutputStream.flush();
        zipOutputStream.close();
        byteArrayOutputStream.close();

    }

    private Map<String, String> getFileUrls() {
        Map<String, String> result = new HashMap<>();
        for (int i = 1; i <= 13; i++) {
            result.put(i + ".jpg", "http://localhost:7070/img/" + i + ".jpg");
        }
        return result;
    }
}
