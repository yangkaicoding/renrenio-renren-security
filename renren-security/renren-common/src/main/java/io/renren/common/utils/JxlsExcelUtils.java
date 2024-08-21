package io.renren.common.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReadStatus;
import org.jxls.reader.XLSReader;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.SAXException;

import java.io.*;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 * Java类作用描述
 * </p>
 *
 * @author: 杨凯
 * @since: 2020/6/8 15:39
 */
public class JxlsExcelUtils {

    private static Logger logger = LoggerFactory.getLogger(JxlsExcelUtils.class);

    public static void exportExcel(InputStream is, OutputStream os, Map<String, Object> model) throws IOException {
        Context context = PoiTransformer.createInitialContext();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        Transformer transformer = jxlsHelper.createTransformer(is, os);
        //获得配置
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
        //设置静默模式，不报警告
        //evaluator.getJexlEngine().setSilent(true);
        //函数强制，自定义功能
        Map<String, Object> funcs = new HashMap<String, Object>();
        funcs.put("utils", new JxlsExcelUtils());
        //添加自定义功能
        evaluator.getJexlEngine().setFunctions(funcs);
        //必须要这个，否者表格函数统计会错乱
        jxlsHelper.setUseFastFormulaProcessor(false).processTemplate(context, transformer);
    }

    public static void exportExcel(File xls, File out, Map<String, Object> model) throws FileNotFoundException, IOException {
        exportExcel(new FileInputStream(xls), new FileOutputStream(out), model);
    }

    public static void exportExcel(String templatePath, OutputStream os, Map<String, Object> model) throws Exception {
        File template = getTemplate(templatePath);
        if (template != null) {
            exportExcel(new FileInputStream(template), os, model);
        } else {
            throw new Exception("Excel 模板未找到。");
        }
    }


    /**
     * @param file        multipartFile对象，前台传过来，
     * @param xmlLocation xml解析表格的映射文件路径
     * @return
     * @throws IOException
     * @throws SAXException
     */
    public static List<?> readExcel(MultipartFile file, String xmlLocation) throws IOException, SAXException {
        try {
            //新建一个list存放对象
            List<?> list = new ArrayList<>();
            //读取xml文件配置获得流
            //InputStream inputXML = new BufferedInputStream(JxlsUtils.class.getClassLoader().getSystemResourceAsStream(xmlLocation)/*getClass().getClassLoader().getResourceAsStream(xmlLocation)*/);
            InputStream inputXML = new BufferedInputStream(JxlsExcelUtils.class.getClassLoader().getResourceAsStream(xmlLocation));
            XLSReader mainReader;
            mainReader = ReaderBuilder.buildFromXML(inputXML);
            //得到excel文件的流
            InputStream inputXLS;
            inputXLS = new BufferedInputStream(file.getInputStream());
            Object entity = new Object();

            Map<String, Object> map = new HashMap<>();

            map.put("entity", entity);
            map.put("list", list);
            XLSReadStatus readStatus;
            //jxls读取excel文件并excel中的数据自动放入list集合中
            //读取的状态
            readStatus = mainReader.read(inputXLS, map);


            //关闭流
            if (inputXML != null) {
                inputXML.close();
            }
            if (inputXLS != null) {
                inputXLS.close();
            }
            if (readStatus.isStatusOK()) {
                return list;
            } else {//读取状态不正确
                return null;
            }
        } catch (IOException e) {
            logger.info("IO异常");
            return null;
        } catch (SAXException e) {
            logger.info("SAX异常");
            return null;
        } catch (InvalidFormatException e) {
            logger.info("InvalidFormat异常");
            return null;
        }
    }

    //获取jxls模版文件
    public static File getTemplate(String path) {
        File template = new File(path);
        if (template.exists()) {
            return template;
        }
        return null;
    }

    // 日期格式化
    public String dateFmt(Date date, String fmt) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(fmt);
            return dateFmt.format(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }


    public String tranferToYuan(long amountPrice) {
        return new BigDecimal(amountPrice).divide(new BigDecimal(100), 2, BigDecimal.ROUND_HALF_UP).toString();
    }


    /**
     * 秒数转字符串
     *
     * @param mills
     * @param format
     * @return
     */
    public String dateFmtByLong(long mills, String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        try {
            if (mills != 0) {
                Date date = new Date((long) mills * 1000L);
                String dateStr = sdf.format(date);
                return dateStr;
            } else {
                return "0";
            }

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            return "0";
        }
    }

    /**
     * 秒数转字符串  yyyy-MM-dd hh:mm:ss
     *
     * @param mills
     * @return
     */
    public String dateFmtByLong(long mills) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        try {
            if (mills != 0) {
                Date date = new Date((long) mills * 1000L);
                String dateStr = sdf.format(date);
                return dateStr;
            } else {
                return "0";
            }

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            return "0";
        }
    }

    // if判断
    public Object ifelse(boolean b, Object o1, Object o2) {
        return b ? o1 : o2;
    }

}
