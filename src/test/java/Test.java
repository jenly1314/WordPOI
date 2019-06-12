import com.king.poi.word.WordPOI;

/**
 * 测试示例
 * @author <a href="mailto:jenly1314@gmail.com">Jenly</a>
 */
public class Test {

    public static void main(String[] args) {
        try {

//            ParseConfig config  = new ParseConfig.Builder().serializable(true).build();
//            WordPOI.wordToEntity("C:/Api1.docx","C:/bean/","com.king.poi.bean",config,"Result","PageInfo");
            WordPOI.wordToEntity(Test.class.getResourceAsStream("Api1.docx"),false,"C:/bean/","com.king.poi.bean","Result","PageInfo");
            WordPOI.wordToEntity(Test.class.getResourceAsStream("Api2.doc"),true,"C:/bean/","com.king.poi.bean","TestBean");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
