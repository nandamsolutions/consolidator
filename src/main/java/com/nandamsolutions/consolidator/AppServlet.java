package com.nandamsolutions.consolidator;

import static com.nandamsolutions.consolidator.WorkbookUtils.workbook;
import static com.nandamsolutions.consolidator.WorkbookUtils.WorkbookType.XLSX;
import static com.nandamsolutions.consolidator.WorkbookUtils.WorkbookType.get;

import java.io.IOException;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.jetty.server.Server;
import org.eclipse.jetty.webapp.WebAppContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.nandamsolutions.consolidator.WorkbookUtils.WorkbookType;

public class AppServlet extends HttpServlet {
    private static final long serialVersionUID = 1296780770721726848L;
    
    private static final Logger LOGGER = LoggerFactory.getLogger(AppServlet.class);
    
    private final ServletFileUpload fileUpload = new ServletFileUpload(new DiskFileItemFactory());
    private final WorkbookService workbookService = new WorkbookService();
    private static final ObjectMapper MAPPER = new ObjectMapper();
    
    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException ,IOException {
        Map<String, Double> totalsData = MAPPER.readValue(req.getParameter("data"), StringDoubleMap.class);
        workbookService.write(totalsData, resp.getOutputStream());
        resp.getOutputStream().flush();
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        Map<String, Map<String, Double>> data = new HashMap<>();
        try {
            fileUpload.parseRequest(req).forEach(item -> {
                try {
                    WorkbookType workbookType = get(item.getName());
                    if (workbookType == null) {
                        LOGGER.error("Unsupported format file: " + item.getName());
                    }
                    Workbook wb = workbook(item.getInputStream(), XLSX.equals(workbookType));
                    data.put(item.getName(), workbookService.read(wb));
                } catch (Exception e) {
                    LOGGER.error("Failed to process file: " + item.getName(), e);
                }
            });
        } catch (FileUploadException e) {
            LOGGER.error("Failed to process uploaded files", e);
        }
        PrintWriter writer = resp.getWriter();
        writer.println(MAPPER.writeValueAsString(data));
        writer.flush();
    }
    
    public static void main(String[] args) throws Exception {
        Server server = new Server(8080);
        WebAppContext webapp = new WebAppContext(AppServlet.class.getClassLoader().getResource("webapp").toExternalForm(), "/");
        webapp.addServlet(AppServlet.class, "/app");
        server.setHandler(webapp);
        server.start();
        server.join();
    }
    
    public static class StringDoubleMap extends LinkedHashMap<String, Double> {
        private static final long serialVersionUID = 5516071111728925858L;
    }
}
