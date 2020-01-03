package me.xueyao;

import me.xueyao.utils.WordUtils;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class SimonItextWordApplicationTests {
    @Test
    void contextLoads () throws Exception {
        WordUtils wordUtils = new WordUtils();
        wordUtils.openDocument("/Users/simonxue/Desktop/12345.doc");
        //wordUtils.insertTitlePattern("小明", RtfParagraphStyle.STYLE_HEADING_1);
        wordUtils.insertTableName("表格", 15, 15, 1);
        wordUtils.insertSimpleTable(3, 1);
        wordUtils.closeDocument();
    }

}
