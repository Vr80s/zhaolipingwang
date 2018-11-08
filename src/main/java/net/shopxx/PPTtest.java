package net.shopxx;


import org.apache.commons.io.FileUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class PPTtest {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();//创建幻灯片
        // 安添加顺序进行填充，背景图为起始添加项，如先添加文本后添加背景图折会覆盖之前添加的文本
        // 幻灯片添加图片
        byte[] bt = FileUtils.readFileToByteArray(new File("C://Users/Administrator/Desktop/找礼品商城文件/u18395.png"));
        XSLFPictureData idx = ppt.addPicture(bt, XSLFPictureData.PictureType.PNG); // 设置添加图片格式
        XSLFPictureShape pic = slide.createPicture(idx); // 添加图片
        pic.setAnchor(new Rectangle(0 , 0, (int) ppt.getPageSize().getWidth(), (int) ppt.getPageSize().getHeight())); // 图片添加位置

        // 幻灯片添加文本
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(10,10, 0, 0));
        textBox.addNewTextParagraph().addNewTextRun().setText("创建幻灯片");

        // 幻灯片保存路径
        ppt.write(new FileOutputStream("D://MyPPT.pptx"));
    }
}
