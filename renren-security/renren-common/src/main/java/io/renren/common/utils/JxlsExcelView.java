package io.renren.common.utils;

import org.springframework.web.servlet.view.AbstractView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Map;

/**
 * <p>
 * Java类作用描述
 * </p>
 *
 * @author: 杨凯
 * @since: 2020/6/8 15:38
 */
public class JxlsExcelView extends AbstractView {

    private String templatePath;
    private String exportFileName;
    private static final String EXCEL_TYPE = "excel_type";
    private static final String CONTENT_TYPE = "application/vnd.ms-excel";

    /**
     * @param templatePath   模版相对于当前classpath路径
     * @param exportFileName 导出文件名
     */
    public JxlsExcelView(String templatePath, String exportFileName) {
        this.templatePath = templatePath;
        if (exportFileName != null) {
            try {
                exportFileName = URLEncoder.encode(exportFileName, "UTF-8");
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
        }
        this.exportFileName = exportFileName;
        setContentType(CONTENT_TYPE);
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request, HttpServletResponse response) throws Exception {
        response.setContentType(getContentType());
        Object fileType = model.get(EXCEL_TYPE);
        response.setHeader("content-disposition", "attachment;filename=" + exportFileName + "." + (null == fileType ? "xls" : String.valueOf(fileType)));
        ServletOutputStream os = response.getOutputStream();
        InputStream is = getClass().getClassLoader().getResourceAsStream(templatePath);
        JxlsExcelUtils.exportExcel(is, os, model);
        is.close();
    }
}
