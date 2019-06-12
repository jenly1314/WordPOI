package com.king.poi.word;

import java.util.List;

/**
 * 实体属性
 * @author <a href="mailto:jenly1314@gmail.com">Jenly</a>
 */
public class EntityProperty {

    private String packageName;

    private String className;

    private List<FieldProperty> filedPropertyList;


    public EntityProperty() {

    }

    public EntityProperty(String packageName, String className) {
        this.packageName = packageName;
        this.className = className;
    }

    public EntityProperty(String packageName, String className, List<FieldProperty> filedPropertyList) {
        this.packageName = packageName;
        this.className = className;
        this.filedPropertyList = filedPropertyList;

    }

    public String getPackageName() {
        return packageName;
    }

    public void setPackageName(String packageName) {
        this.packageName = packageName;
    }

    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }

    public List<FieldProperty> getFiledPropertyList() {
        return filedPropertyList;
    }

    public void setFiledPropertyList(List<FieldProperty> filedPropertyList) {
        this.filedPropertyList = filedPropertyList;
    }

    @Override
    public String toString() {
        return "EntityProperty{" +
                "packageName='" + packageName + '\'' +
                ", className='" + className + '\'' +
                ", filedPropertyList=" + filedPropertyList +
                '}';
    }
}
