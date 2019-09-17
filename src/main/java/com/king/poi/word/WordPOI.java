package com.king.poi.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;


/**
 * 核心功能：将文档中表格定义的实体转换成Java实体对象
 * <br/> 1. 支持解析doc格式和docx格式的Word文档
 * <br/> 2. 支持批量解析Word文档并转换成实体
 * <br/> 3. 解析配置支持自定义，详情请查看{@link ParseConfig}相关配置
 * <br/> 4. 虽然解析可配置，但因文档内容的不可控，解析转换也具有一定的局限性
 *
 * @author <a href="mailto:jenly1314@gmail.com">Jenly</a>
 */
public final class WordPOI {

    private WordPOI(){
        throw new AssertionError();
    }

    /**
     * 通过流得到XWPFDocument (.docx格式的文档)
     * @param inputStream
     * @return
     * @throws IOException
     */
    private static XWPFDocument getXWPFDocument(InputStream inputStream) throws IOException {
        return new XWPFDocument(inputStream);
    }

    /**
     * 通过流得到HWPFDocument (.doc格式的文档)
     * @param inputStream
     * @return
     * @throws IOException
     */
    private static HWPFDocument getHWPFDocument(InputStream inputStream) throws IOException {
        return new HWPFDocument(inputStream);
    }

    /**
     * 获取TableIterator
     * @param inputStream
     * @return
     * @throws IOException
     */
    private static TableIterator getTableIterator(InputStream inputStream) throws IOException {
        return new TableIterator(getHWPFDocument(inputStream).getRange());
    }

    /**
     * 获取XWPFTable列表
     * @param inputStream
     * @return
     * @throws IOException
     */
    private static List<XWPFTable> getXWPFTables(InputStream inputStream) throws IOException {
        return getXWPFDocument(inputStream).getTables();
    }

    /**
     * Word文档转实体
     * @param wordPath 文档路径
     * @param outPath 输出路径
     * @param packageName 实体类包名
     * @param entityNames 实体类名称
     * @throws Exception
     */
    public static void wordToEntity(String wordPath,String outPath,String packageName,String... entityNames) throws Exception{
        wordToEntity(wordPath,outPath,packageName,new ParseConfig(),entityNames);
    }

    /**
     * Word文档转实体
     * @param wordPath 文档路径
     * @param outPath 输出路径
     * @param packageName 实体类包名
     * @param parseConfig 解析配置
     * @param entityNames 实体类名称
     * @throws Exception
     */
    public static void wordToEntity(String wordPath,String outPath,String packageName,ParseConfig parseConfig,String... entityNames) throws Exception{
        if(wordPath!=null){
            boolean isDoc = wordPath.endsWith("doc");
            wordToEntity(new FileInputStream(wordPath),isDoc,outPath,packageName,parseConfig,entityNames);
        }

    }

    /**
     * Word文档转实体
     * @param inputStream
     * @param isDoc  true 表示是".doc"格式的文件，false 表示是“.docx”格式的文件
     * @param outPath 输出路径
     * @param packageName 实体类包名
     * @param entityNames 实体名称
     * @throws IOException
     */
    public static void wordToEntity(InputStream inputStream,boolean isDoc,String outPath,String packageName,String... entityNames) throws IOException{
        wordToEntity(inputStream,isDoc,outPath,packageName,new ParseConfig(),entityNames);
    }

    /**
     * Word文档转实体
     * @param inputStream 需要解析的文档的流
     * @param isDoc true 表示是".doc"格式的文件，false 表示是“.docx”格式的文件
     * @param outPath 输出路径
     * @param packageName 实体类包名
     * @param parseConfig 解析配置
     * @param entityNames 实体名称
     * @throws IOException
     */
    public static void wordToEntity(InputStream inputStream,boolean isDoc,String outPath,String packageName,ParseConfig parseConfig,String... entityNames) throws IOException{
        if(isDoc){
            docToEntity(inputStream,outPath,packageName,parseConfig,entityNames);
        }else{
            docxToEntity(inputStream,outPath,packageName,parseConfig,entityNames);
        }
    }

    /**
     * docx格式的文档转实体
     * @param inputStream 需要解析的文档的流
     * @param outPath 输出路径
     * @param packageName 实体类包名
     * @param parseConfig 解析配置
     * @param entityNames 实体名称
     * @throws IOException
     */
    private static void docxToEntity(InputStream inputStream,String outPath,String packageName,ParseConfig parseConfig,String... entityNames) throws IOException{

        boolean isParseEntityName = parseConfig.isParseEntityName();
        int length = entityNames.length;
        if(isParseEntityName || length>0){
            List<XWPFTable> xwpfTableList = getXWPFTables(inputStream);

            int size = xwpfTableList.size();
            //遍历文档所有表格
            for (int i = parseConfig.getStartTable(); i < size; i++) {
                if(i > length && !isParseEntityName){
                    return;
                }
                //得到文档当前表格
                XWPFTable table = xwpfTableList.get(i);
                int rows = table.getNumberOfRows();

                String entityName;
                if(isParseEntityName){
                    entityName = table.getRow(parseConfig.getEntityNameRow()).getCell(parseConfig.getEntityNameColumn()).getText();
                }else{
                    entityName = entityNames[i];
                }
                EntityProperty entityProperty = new EntityProperty(packageName,entityName);

                Map<String,String> transformations = parseConfig.getTransformations();
                boolean isTransfer = transformations != null && !transformations.isEmpty();

                List<FieldProperty> filedPropertyList = new ArrayList<>();
                //遍历文档当前表格的当前行
                for (int j = parseConfig.getStartRow(); j < rows; j++){
                    XWPFTableRow row = table.getRow(j);
                    List<XWPFTableCell> cells = row.getTableCells();
                    int cols = cells.size();

                    FieldProperty filedProperty = new FieldProperty();
                    //遍历文档当前表格当前行的当前列
                    for (int k = parseConfig.getStartColumn(); k < cols; k++) {
                        XWPFTableCell cell = cells.get(k);
                        if(k == parseConfig.getFieldNameColumn()){
                            filedProperty.setName(cell.getText());
                        }else if(k == parseConfig.getFieldTypeColumn()){
                            String type = cell.getText();
                            if(isTransfer && transformations.containsKey(type)){
                                filedProperty.setType(transformations.get(type));
                            }else{
                                filedProperty.setType(type);
                            }
                        }else if(k == parseConfig.getFieldDescColumn()){
                            filedProperty.setDesc(cell.getText());
                        }
                    }
                    String name = filedProperty.getName();
                    //过滤掉为空的字段名称
                    if(name != null && name.trim().length() > 0){
                        filedPropertyList.add(filedProperty);
                    }
                }

                entityProperty.setFiledPropertyList(filedPropertyList);

                generateEntity(entityProperty,false,outPath,parseConfig);

            }

        }

    }

    /**
     * doc格式的文档转实体
     * @param inputStream 需要解析的文档的流
     * @param outPath 输出路径
     * @param packageName 实体类包名
     * @param parseConfig 解析配置
     * @param entityNames 实体名称
     * @throws IOException
     */
    private static void docToEntity(InputStream inputStream,String outPath,String packageName,ParseConfig parseConfig,String... entityNames) throws IOException{

        boolean isParseEntityName = parseConfig.isParseEntityName();
        int length = entityNames.length;
        if(isParseEntityName || length>0){
            TableIterator tableIterator = getTableIterator(inputStream);
            int index = 0;

            for(int i = 0; i<parseConfig.getStartTable(); i++){
                tableIterator.hasNext();
            }

            Map<String,String> transformations = parseConfig.getTransformations();
            boolean isTransfer = transformations != null && !transformations.isEmpty();
            while (tableIterator.hasNext()){

                if(index > length && !isParseEntityName){
                    return;
                }

                Table table = tableIterator.next();
                int rows = table.numRows();

                String entityName;
                if(isParseEntityName){
                    entityName = table.getRow(parseConfig.getEntityNameRow()).getCell(parseConfig.getEntityNameColumn()).text().replace("\u0007","");
                }else{
                    entityName =  entityNames[index];
                }
                EntityProperty entityProperty = new EntityProperty(packageName,entityName);

                List<FieldProperty> filedPropertyList = new ArrayList<>();
                //遍历文档当前表格的当前行
                for (int j = parseConfig.getStartRow(); j < rows; j++){
                    TableRow row = table.getRow(j);
                    int cols = row.numCells();
                    FieldProperty filedProperty = new FieldProperty();
                    //遍历文档当前表格当前行的当前列
                    for (int k = parseConfig.getStartColumn(); k < cols; k++) {
                        TableCell cell = row.getCell(k);
                        if(k == parseConfig.getFieldNameColumn()){
                            filedProperty.setName(cell.getParagraph(0).text().replace("\u0007",""));
                        }else if(k == parseConfig.getFieldTypeColumn()){
                            String type = cell.getParagraph(0).text().replace("\u0007","");
                            if(isTransfer && transformations.containsKey(type)){
                                filedProperty.setType(transformations.get(type));
                            }else{
                                filedProperty.setType(type);
                            }
                        }else if(k == parseConfig.getFieldDescColumn()){
                            filedProperty.setDesc(cell.getParagraph(0).text().replace("\u0007",""));
                        }
                    }

                    String name = filedProperty.getName();
                    //过滤掉为空的字段名称
                    if(name != null && name.trim().length() > 0){
                        filedPropertyList.add(filedProperty);
                    }
                }

                entityProperty.setFiledPropertyList(filedPropertyList);

                generateEntity(entityProperty,true,outPath,parseConfig);

                index++;
            }

        }

    }


    /**
     * 生成实体对象
     * @param entityProperty 实体属性
     * @param isDoc true 表示是".doc"格式的文件，false 表示是“.docx”格式的文件
     * @param outPath 输出路径
     * @param parseConfig 解析配置
     * @throws IOException
     */
    private static void generateEntity(EntityProperty entityProperty,boolean isDoc,String outPath,ParseConfig parseConfig) throws IOException{
        String className = entityProperty.getClassName();
        String javaName =  className + ".java";

        StringBuffer buffer = new StringBuffer();

        buffer.append("/* ")
                .append("<a href=\"https://github.com/jenly1314/WordPOI\">WordPOI</a>")
                .append(" */\n");

        buffer.append("package "+ entityProperty.getPackageName()+";\n\n");
        //是否使用Lobmok
        boolean useLombok = parseConfig.isUseLombok();
        //是否生成get和set方法
        boolean genGetterAndSetter = parseConfig.isGenGetterAndSetter();
        //是否生成toString方法
        boolean genToString = parseConfig.isGenToString();

        //是否包含泛型
        boolean genericity = false;
        //是否导入List
        boolean importList = false;
        //是否导入Set
        boolean importSet = false;
        //是否导入Map
        boolean importMap = false;

        List<FieldProperty> filedPropertyList = entityProperty.getFiledPropertyList();
        if(filedPropertyList !=null ){

            StringBuffer bufferField = new StringBuffer();
            StringBuffer bufferGetSet = new StringBuffer();
            StringBuffer bufferToString = new StringBuffer();

            //toString开头
            if(genToString && !useLombok){
                bufferToString.append("\t@Override\n")
                        .append("\tpublic String toString(){\n")
                        .append("\t\treturn ")
                        .append("\"").append(className).append("{\" + \n");
            }

            int index = 0;
            int size =  filedPropertyList.size();
            for(FieldProperty filedProperty : filedPropertyList){
                bufferField.append("\t/** " + filedProperty.getDesc() + " */\n");
                bufferField.append("\tprivate " + filedProperty.getType() + " " + filedProperty.getName() + ";\n");

                String fieldType = filedProperty.getType();
                if(fieldType!=null){
                    //遍历是否包含泛型
                    if(fieldType.contains("<T>") || fieldType.equals("T")){
                        genericity = true;
                    }

                    //遍历是否包含集合
                    if(fieldType.startsWith("List")){
                        importList = true;
                    }else if(fieldType.startsWith("Set")){
                        importSet = true;
                    }else if(fieldType.startsWith("Map")){
                        importMap = true;
                    }

                }

                if(!useLombok){//只有当不使用Lombok时，才会根据genGetterAndSetter 和 genToString 来判断是否生成get、set、toString等方法

                    String fieldName = filedProperty.getName();
                    //是否生成get和set方法
                    if(genGetterAndSetter){
                        String getterName = fieldName.substring(0,1).toUpperCase() + fieldName.substring(1);

                        //Generate getter
                        if("boolean".equalsIgnoreCase(fieldType)){
                            bufferGetSet.append("\tpublic " + fieldType + " is" + getterName + "(){\n")
                                    .append("\t\treturn " + fieldName + ";\n")
                                    .append("\t}\n\n");
                        }else{
                            bufferGetSet.append("\tpublic " + fieldType + " get" + getterName + "(){\n")
                                    .append("\t\treturn " + fieldName + ";\n")
                                    .append("\t}\n\n");
                        }
                        //Generate setter
                        bufferGetSet.append("\tpublic void set" + getterName + "(" + fieldType + " " + filedProperty.getName() + "){\n")
                                .append("\t\tthis." + fieldName + " = " + fieldName + ";\n")
                                .append("\t}\n\n");
                    }

                    if(genToString){//生成toString方法
                        bufferToString.append("\t\t\t\t\"").append(fieldName).append("=\" + ")
                                .append(fieldName);
                        if(index < size-1){
                            bufferToString.append(" + \",\" + \n");
                        }else{
                            bufferToString.append(" + \"}\";\n");
                        }
                    }
                }

                index++;
            }

            if(parseConfig.isSerializable()){//导入Serializable
                buffer.append("import java.io.Serializable;\n");
            }

            if(importList){//导入List
                buffer.append("import java.util.List;\n");
            }

            if(importSet){//导入Set
                buffer.append("import java.util.Set;\n");
            }

            if(importMap){//导入Map
                buffer.append("import java.util.Map;\n");
            }

            if(useLombok){//导入Lombok
                buffer.append("import lombok.Data;\n");
            }
            buffer.append("\n");

            //是否显示头部注释说明
            if(parseConfig.isShowHeader()){
                String header = parseConfig.getHeader();
                if(header!=null && header.length()>0){
                    buffer.append("/**\n");
                    buffer.append(" * ")
                            .append(header.replace("\n","\n * "));
                    buffer.append("\n */\n");
                }
            }

            if(useLombok){//添加Lombok的data注解
                buffer.append("@Data\n");
            }

            //拼接类
            buffer.append("public class " + className);
            if(genericity){
                buffer.append("<T>");
            }

            //序列化
            if(parseConfig.isSerializable()){
                buffer.append(" implements Serializable");
            }

            buffer.append(" {\n\n");

            //添加字段
            buffer.append(bufferField).append("\n");

            if(!useLombok){

                if(genGetterAndSetter){
                    buffer.append(bufferGetSet);
                }

                if(genToString){
                    //toString结尾
                    bufferToString.append("\t}\n\n");
                    buffer.append(bufferToString);
                }
            }

        }

        buffer.append("\n}\n\n");

        File dirFile = new File(outPath);
        if(dirFile!=null && !dirFile.exists()){
            dirFile.mkdirs();
        }

        //将内容写入到文件
        File file = new File(outPath,javaName);
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file),parseConfig.getCharsetName()));
        //doc格式的文档解析后会有\u0007字符需统一过滤
        writer.write(isDoc ? buffer.toString().replace("\u0007","") : buffer.toString());
        writer.flush();
        writer.close();

    }

}
