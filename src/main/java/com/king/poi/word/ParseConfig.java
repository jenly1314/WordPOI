package com.king.poi.word;

/**
 * 解析配置
 * @author <a href="mailto:jenly1314@gmail.com">Jenly</a>
 */
public class ParseConfig {

    /**
     * 开始表格，默认从0开始
     */
    private int startTable = 0;

    /**
     * 开始行，默认：从1开始，因为第0行一般是标题说明
     */
    private int startRow = 1;
    /**
     * 开始列，默认从0开始
     */
    private int startColumn = 0;
    /**
     * 字段名称所在列，默认：第0列
     */
    private int fieldNameColumn = 0;
    /**
     * 字段类型所在列，默认：第1列
     */
    private int fieldTypeColumn = 1;
    /**
     * 字段注释说明所在列，默认：第2列
     */
    private int fieldDescColumn = 2;
    /**
     * 字符集编码，默认：UTF-8
     */
    private String charsetName = "UTF-8";

    /**
     * 是否生成get和set方法
     */
    private boolean genGetterAndSetter = true;

    /**
     * 是否生成toString方法
     */
    private boolean genToString = true;

    /**
     * 是否使用Lombok
     */
    private boolean useLombok;

    /**
     * 是否实现Serializable序列化
     */
    private boolean serializable;

    /**
     * 是否显示头注释
     */
    private boolean showHeader = true;

    /**
     * 头注释内容
     */
    private String header = "Created by WordPOI";

    public ParseConfig() {
    }

    private ParseConfig(Builder builder) {
        startTable = builder.startTable;
        startRow = builder.startRow;
        startColumn = builder.startColumn;
        fieldNameColumn = builder.fieldNameColumn;
        fieldTypeColumn = builder.fieldTypeColumn;
        fieldDescColumn = builder.fieldDescColumn;
        charsetName = builder.charsetName;
        genGetterAndSetter = builder.genGetterAndSetter;
        genToString = builder.genToString;
        useLombok = builder.useLombok;
        serializable = builder.serializable;
        showHeader = builder.showHeader;
        header = builder.header;
    }


    public int getStartTable() {
        return startTable;
    }

    public int getStartRow() {
        return startRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public int getFieldNameColumn() {
        return fieldNameColumn;
    }

    public int getFieldTypeColumn() {
        return fieldTypeColumn;
    }

    public int getFieldDescColumn() {
        return fieldDescColumn;
    }

    public String getCharsetName() {
        return charsetName;
    }

    public boolean isGenGetterAndSetter() {
        return genGetterAndSetter;
    }

    public boolean isGenToString() {
        return genToString;
    }

    public boolean isUseLombok() {
        return useLombok;
    }

    public boolean isSerializable() {
        return serializable;
    }

    public boolean isShowHeader() {
        return showHeader;
    }

    public String getHeader() {
        return header;
    }


    public static final class Builder {
        /**
         * 开始表格，默认从0开始
         */
        private int startTable = 0;

        /**
         * 开始行，默认：从1开始，因为第0行一般是标题说明
         */
        private int startRow = 1;
        /**
         * 开始列，默认从0开始
         */
        private int startColumn = 0;
        /**
         * 字段名称所在列，默认：第0列
         */
        private int fieldNameColumn = 0;
        /**
         * 字段类型所在列，默认：第1列
         */
        private int fieldTypeColumn = 1;
        /**
         * 字段注释说明所在列，默认：第2列
         */
        private int fieldDescColumn = 2;
        /**
         * 字符集编码，默认：UTF-8
         */
        private String charsetName = "UTF-8";

        /**
         * 是否生成get和set方法
         */
        private boolean genGetterAndSetter = true;

        /**
         * 是否生成toString方法
         */
        private boolean genToString = true;

        /**
         * 是否使用Lombok
         */
        private boolean useLombok;

        /**
         * 是否实现Serializable序列化
         */
        private boolean serializable;

        /**
         * 是否显示头注释
         */
        private boolean showHeader = true;

        /**
         * 头注释内容
         */
        private String header = "Created by WordPOI";

        public Builder() {
        }

        /**
         * 设置开始表格
         * @param startTable {@link ParseConfig#startTable}
         * @return {@link Builder}
         */
        public Builder startTable(int startTable) {
            this.startTable = startTable;
            return this;
        }

        /**
         * 设置开始行
         * @param startRow {@link ParseConfig#startRow}
         * @return {@link Builder}
         */
        public Builder startRow(int startRow) {
            this.startRow = startRow;
            return this;
        }

        /**
         * 设置开始列
         * @param startColumn {@link ParseConfig#startColumn}
         * @return {@link Builder}
         */
        public Builder startColumn(int startColumn) {
            this.startColumn = startColumn;
            return this;
        }

        /**
         * 设置字段名称所在列
         * @param fieldNameColumn {@link ParseConfig#fieldNameColumn}
         * @return {@link Builder}
         */
        public Builder fieldNameColumn(int fieldNameColumn) {
            this.fieldNameColumn = fieldNameColumn;
            return this;
        }

        /**
         * 设置字段类型所在列
         * @param fieldTypeColumn {@link ParseConfig#fieldTypeColumn}
         * @return {@link Builder}
         */
        public Builder fieldTypeColumn(int fieldTypeColumn) {
            this.fieldTypeColumn = fieldTypeColumn;
            return this;
        }

        /**
         * 设置字段注释描述所在列
         * @param fieldDescColumn {@link ParseConfig#fieldDescColumn}
         * @return {@link Builder}
         */
        public Builder fieldDescColumn(int fieldDescColumn) {
            this.fieldDescColumn = fieldDescColumn;
            return this;
        }

        /**
         * 设置字符集编码
         * @param charsetName {@link ParseConfig#charsetName}
         * @return {@link Builder}
         */
        public Builder charsetName(String charsetName) {
            this.charsetName = charsetName;
            return this;
        }

        /**
         * 设置是否生成get和set方法
         * @param genGetterAndSetter {@link ParseConfig#genGetterAndSetter}
         * @return {@link Builder}
         */
        public Builder genGetterAndSetter(boolean genGetterAndSetter) {
            this.genGetterAndSetter = genGetterAndSetter;
            return this;
        }

        /**
         * 设置是否生成toString方法
         * @param genToString {@link ParseConfig#genToString}
         * @return {@link Builder}
         */
        public Builder genToString(boolean genToString) {
            this.genToString = genToString;
            return this;
        }

        /**
         * 设置是否使用Lombok
         * @param useLombok {@link ParseConfig#useLombok}
         * @return {@link Builder}
         */
        public Builder useLombok(boolean useLombok) {
            this.useLombok = useLombok;
            return this;
        }

        /**
         * 设置是否序列化
         * @param serializable {@link ParseConfig#serializable}
         * @return {@link Builder}
         */
        public Builder serializable(boolean serializable) {
            this.serializable = serializable;
            return this;
        }


        /**
         * 设置是否显示 class头注释
         * @param showHeader {@link ParseConfig#showHeader}
         * @return {@link Builder}
         */
        public Builder showHeader(boolean showHeader) {
            this.showHeader = showHeader;
            return this;
        }

        /**
         * 设置class头注释内容
         * @param header {@link ParseConfig#header}
         * @return {@link Builder}
         */
        public Builder header(String header) {
            this.header = header;
            return this;
        }

        /**
         * 创建ParseConfig
         * @return {@link ParseConfig}
         */
        public ParseConfig build() {
            return new ParseConfig(this);
        }
    }
}
