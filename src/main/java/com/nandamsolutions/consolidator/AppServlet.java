package com.nandamsolutions.consolidator;

import static com.nandamsolutions.consolidator.WorkbookUtils.workbook;
import static com.nandamsolutions.consolidator.WorkbookUtils.WorkbookType.XLSX;
import static com.nandamsolutions.consolidator.WorkbookUtils.WorkbookType.get;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.stream.Collectors;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.jetty.server.Server;
import org.eclipse.jetty.webapp.WebAppContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.nandamsolutions.consolidator.WorkbookUtils.WorkbookType;

public class AppServlet extends HttpServlet {
    private static final long serialVersionUID = 1296780770721726848L;

    private static final Logger LOGGER = LoggerFactory.getLogger(AppServlet.class);

    private final ServletFileUpload fileUpload = new ServletFileUpload(new DiskFileItemFactory());
    private final WorkbookService workbookService = new WorkbookService();

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        try {
            List<FileItem> items = fileUpload.parseRequest(req);
            String filename = new String(items.stream().filter(item -> item.getFieldName().equals("filename")).findAny().get().get());
            String heading = new String(items.stream().filter(item -> item.getFieldName().equals("heading")).findAny().get().get());
            List<Workbook> workbooks = items.stream().filter(item -> item.getName()!=null).map(item -> {
                try {
                    WorkbookType workbookType = get(item.getName());
                    if (workbookType == null) {
                        LOGGER.error("Unsupported format file: " + item.getName());
                    }
                    return workbook(item.getInputStream(), XLSX.equals(workbookType));
                } catch (Exception e) {
                    return null;
                }
            }).collect(Collectors.toList());
            try (Workbook consolidated = workbookService.createCopy(workbooks.get(0), heading)) {
                workbookService.mergeBooks(workbooks, consolidated);
                resp.setHeader("Content-Disposition", "attachment; filename=\""+filename+".xls\"");
                consolidated.write(resp.getOutputStream());
            }
        } catch (FileUploadException e) {
            LOGGER.error("Failed to process uploaded files", e);
        }
    }

    public static void main(String[] args) throws Exception {
        Server server = new Server(8080);
        WebAppContext webapp = new WebAppContext(
                AppServlet.class.getClassLoader().getResource("webapp").toExternalForm(), "/");
        webapp.addServlet(AppServlet.class, "/app");
        server.setHandler(webapp);
        server.start();
        server.join();
    }

    public static class StringDoubleMap extends LinkedHashMap<String, Double> {
        private static final long serialVersionUID = 5516071111728925858L;
    }
}
