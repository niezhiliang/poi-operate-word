package cn.isuyu.poi.operate.word.controller;

import cn.isuyu.poi.operate.word.dtos.OperateDTO;
import cn.isuyu.poi.operate.word.utils.WordUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Collections;

/**
 * @Author NieZhiLiang
 * @Email nzlsgg@163.com
 * @GitHub https://github.com/niezhiliang
 * @Date 2020/5/7 下午4:50
 */
@RestController
public class IndexController {

    @RequestMapping(value = "operate")
    public String operate(@RequestBody OperateDTO dto) throws Exception {

        //排序 防止替换前面的列后 后面的列索引变换
        Collections.sort(dto.getDynamicRows());
        XWPFDocument document = new XWPFDocument(new FileInputStream("./data/test.docx"));
        WordUtils.dynamicTableReplace(document,dto.getDynamicRows());
        document.write(new FileOutputStream("./data/finish.docx"));

        return "success";
    }

}
