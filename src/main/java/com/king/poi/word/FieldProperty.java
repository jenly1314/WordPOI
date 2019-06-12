package com.king.poi.word;

/**
 * 字段属性
 * @author <a href="mailto:jenly1314@gmail.com">Jenly</a>
 */
public class FieldProperty {

    /**
     * 字段名称
     */
    private String name;

    /**
     * 字段类型
     */
    private String type;
    /**
     * 注释说明
     */
    private String desc;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    @Override
    public String toString() {
        return "FieldInfo{" +
                "name='" + name + '\'' +
                ", type=" + type +
                ", desc='" + desc + '\'' +
                '}';

    }
}
