package com.deepoove.poi.tl.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

@DisplayName("Picture Render test case")
public class PictureRenderTest {

    BufferedImage bufferImage;

    @BeforeEach
    public void init() {
        bufferImage = BytePictureUtils.newBufferImage(100, 100);
        Graphics2D g = (Graphics2D) bufferImage.getGraphics();
        g.setColor(Color.CYAN);
        g.fillRect(0, 0, 100, 100);
        g.setColor(Color.BLACK);
        g.drawString("Java Image", 0, 50);
        g.dispose();
        bufferImage.flush();
    }

    @SuppressWarnings("serial")
    @Test
    public void testPictureRender() throws Exception {
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                // 本地图片
                put("localPicture", new PictureRenderData(120, 120, "src/test/resources/sayi.png"));

                // 图片流文件
                put("localBytePicture",
                        new PictureRenderData(100, 120, ".png", new FileInputStream("src/test/resources/logo.png")));

                // 网络图片
                put("urlPicture", new PictureRenderData(100, 100, ".png",
                        BytePictureUtils.getUrlBufferedImage("http://deepoove.com/images/icecream.png")));

                // java 图片
                put("bufferImagePicture", new PictureRenderData(100, 120, ".png", bufferImage));

                // 不存在图片使用alt文字代替
                PictureRenderData pictureRenderData = new PictureRenderData(120, 120, "src/test/resources/sayi11.png");
                pictureRenderData.setAltMeta("图片不存在");
                put("image", pictureRenderData);
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/render_picture.docx").render(datas);

        FileOutputStream out = new FileOutputStream("out_render_picture.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }

}
