package me.xueyao.controller;

import me.xueyao.utils.WordUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author Simon.Xue
 * @date 2019-12-06 16:45
 **/
@RequestMapping("/test")
@RestController
public class TestController {


    @GetMapping("/test")
    public String test() throws Exception {
        WordUtils wordUtils = new WordUtils();
        wordUtils.openDocument("/Users/simonxue/Desktop/12345.doc");
        //wordUtils.insertTitlePattern("小明", RtfParagraphStyle.STYLE_HEADING_1);
        wordUtils.insertTableName("表格", 15, 15, 1);
        wordUtils.insertSimpleTable(3, 1);
        wordUtils.closeDocument();
        return "";
    }

}
