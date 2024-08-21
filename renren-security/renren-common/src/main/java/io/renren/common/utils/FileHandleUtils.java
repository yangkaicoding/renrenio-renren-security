package io.renren.common.utils;

import cn.hutool.json.JSONUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Date;
import java.util.Map;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;

/**
 * <p>
 * 文件处理
 * </p>
 *
 * @author: 杨凯
 * @since: 2020/6/16 14:07
 */
public class FileHandleUtils {

    @Value("${system.constant.filePath}")
    private String uploadPath;

    private static Logger log = LoggerFactory.getLogger(FileHandleUtils.class);

    /**
     * 生成UUID
     *
     * @return java.lang.String
     * @throws
     * @author YK
     * @date 2024/8/21 下午2:48
     */
    public static String getUUID() {
        String str = UUID.randomUUID().toString();
        String uuidStr = str.replace("-", "");
        return uuidStr;
    }

    /**
     * 判断文件是否为图片
     *
     * @param file
     * @return boolean
     * @throws
     * @author YK
     * @date 2024/8/21 下午2:46
     */
    private boolean isImage(MultipartFile file) {
        boolean flag = false;
        try {
            InputStream is = file.getInputStream();
            BufferedImage bi = ImageIO.read(is);
            if (null == bi) {
                return flag;
            }
            is.close();
            flag = true;
        } catch (Exception e) {
            log.error("判断文件是否为图片错误：{}", e);
        }
        return flag;
    }

    /**
     * 文件上传
     *
     * @param file
     * @return void
     * @throws
     * @author YK
     * @date 2024/8/21 下午2:43
     */
    public void upload(MultipartFile file) {

        log.info("文件名称，{}", file.getOriginalFilename());

        Map<String, Object> fileMap = new ConcurrentHashMap<>();
        // 上传文件存储路径
        String filePath = uploadPath;
        // 设置上传文件时间
        String uploadDate = DateUtils.format(new Date(), "yyyyMMdd");

        // 判断文件是否为图片
        String fileType = null;
        if (!isImage(file)) {
            fileType = "file";
        } else {
            fileType = "image";
        }
        filePath = filePath + "/" + fileType + "/" + uploadDate;
        String fileName = file.getOriginalFilename();

        String oldFileName = fileName;
        // 生成新的文件后缀
        String suffixName = fileName.substring(fileName.lastIndexOf("."), fileName.length());
        // 生成新的文件名称
        fileName = getUUID() + suffixName;
        fileMap.put("fileName", oldFileName);

        File localFile = new File(filePath);
        if (!localFile.exists()) {
            boolean isCreated = localFile.mkdirs();
            if (!isCreated) {
                // 目标上传目录创建失败
                log.info("文件目录创建失败");
            }
        }
        localFile = new File(filePath + "/" + fileName);
        try {
            file.transferTo(localFile);
            fileMap.put("filePath", fileType + "/" + uploadDate + "/" + fileName);
            log.info(JSONUtil.toJsonStr(fileMap));
        } catch (Exception e) {
            log.error("文件上传异常" + e);
            log.error("该文件上传异常：" + oldFileName);
        }
    }

    /**
     * 文件下载
     *
     * @param filePath
     * @param fileName
     * @param res
     * @return void
     * @throws
     * @author YK
     * @date 2024/8/21 下午2:43
     */
    public static void downLoad(String filePath, String fileName, HttpServletResponse res) {
        File file = new File(filePath);
        if (file.exists()) {
            OutputStream os = null;
            BufferedInputStream bis = null;
            try {
                os = res.getOutputStream();
                bis = new BufferedInputStream(new FileInputStream(new File(filePath)));
                res.setContentType("application/octet-stream");
                res.setHeader("Content-Length", String.valueOf(bis.available()));
                res.setHeader("Content-Disposition", "attachment;filename=" + fileName);
                byte[] buff = new byte[1024];

                int i = bis.read(buff);
                while (i != -1) {
                    os.write(buff, 0, buff.length);
                    os.flush();
                    i = bis.read(buff);
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                if (bis != null) {
                    try {
                        bis.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        } else {
            log.info("文件不存在,请检查文件路径是否正确");
        }
    }
}
