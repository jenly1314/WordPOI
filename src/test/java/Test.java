import com.king.poi.word.ParseConfig;
import com.king.poi.word.WordPOI;

/**
 * 测试示例
 * @author <a href="mailto:jenly1314@gmail.com">Jenly</a>
 */
public class Test {

    private static String[] names = {""};

    /**
     * 1、支持解析doc格式和docx格式的Word文档
     * 2、支持批量解析Word文档并转换成实体
     * 3、解析配置支持自定义，详情请查看{@link ParseConfig}相关配置
     * 4、虽然解析可配置，但因文档内容的不可控，解析转换也具有一定的局限性
     */
    public static void main(String[] args) {
        try {

            /**
             * 解析文档中的表格实体，表格包含了实体名称，只需配置 {@link ParseConfig#parseEntityName} 为 true 和相关对应行，即可开启自动解析实体名称，自动解析实体名称
             * {@link ParseConfig}中包含解析时需要的各种配置，方便灵活的支持文档中更多的表格样式
             */
            ParseConfig config  = new ParseConfig.Builder().startRow(2).parseEntityName(true).build();
            WordPOI.wordToEntity(Test.class.getResourceAsStream("Api3.docx"),false,"C:/bean/","com.king.poi.bean",config);
            //解析文档docx格式  需要传生成的对象实体名称
//            WordPOI.wordToEntity(Test.class.getResourceAsStream("Api1.docx"),false,"C:/bean/","com.king.poi.bean","Result","PageInfo");
            //解析文档docx格式  需要传生成的对象实体名称
//            WordPOI.wordToEntity(Test.class.getResourceAsStream("Api2.doc"),true,"C:/bean/","com.king.poi.bean","TestBean");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
