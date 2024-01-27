# WordPOI

[![Download](https://img.shields.io/badge/download-jar-blue.svg)](https://raw.githubusercontent.com/jenly1314/WordPOI/master/libs/word-poi-1.0.1.jar)
[![JitPack](https://jitpack.io/v/jenly1314/WordPOI.svg)](https://jitpack.io/#jenly1314/WordPOI)
[![CI](https://travis-ci.org/jenly1314/WordPOI.svg?branch=master)](https://travis-ci.org/jenly1314/WordPOI)
[![CircleCI](https://circleci.com/gh/jenly1314/WordPOI.svg?style=svg)](https://circleci.com/gh/jenly1314/WordPOI)
[![License](https://img.shields.io/badge/license-Apche%202.0-blue.svg)](http://www.apache.org/licenses/LICENSE-2.0)

WordPOI是一个将Word接口文档转换成JavaBean的工具库，主要目的是减少部分无脑的开发工作。

> 核心功能：将文档中表格定义的实体转换成Java实体对象
 
## 引入

### Gradle:

1. 在Project的 **build.gradle** 或 **setting.gradle** 中添加远程仓库

    ```gradle
    repositories {
        //...
        mavenCentral()
        maven { url 'https://jitpack.io' }
    }
    ```

2. 在Module的 **build.gradle** 里面添加引入依赖项

    ```gradle
    implementation 'com.github.jenly1314:WordPOI:1.0.1'
    ```

* WordPOI依赖的库：
```gradle
implementation 'org.apache.poi:poi:4.1.0'
implementation 'org.apache.poi:poi-ooxml:4.1.0'
implementation 'org.apache.poi:poi-scratchpad:4.1.0'
```

### 本地依赖jar

你也可以通过本地依赖的方式来引入jar库，直接点击左上角的Download[下载最新的jar库](https://raw.githubusercontent.com/jenly1314/WordPOI/master/libs/word-poi-1.0.1.jar)，然后引入到你的工程即可。

## 使用

## WordPOI特性说明  
 1. 支持解析doc格式和docx格式的Word文档
 2. 支持批量解析Word文档并转换成实体
 3. 解析配置支持自定义，详情请查看{@link ParseConfig}相关配置
 4. 虽然解析可配置，但因文档内容的不可控，解析转换也具有一定的局限性

> 只要在文档上定义实体对象时，尽量满足示例文档的规则，就可以规避解析转换时的局限性。

### ParseConfig属性说明
| 属性 | 值类型 | 默认值 | 说明 |
| :------| :------ | :------ | :------ |
| startTable | int |0| 开始表格 |
| startRow | int |1| 开始行 |
| startColumn | int |0| 开始列 |
| fieldNameColumn | int | 0 | 字段名称所在列 |
| fieldTypeColumn | int | 1 | 字段类型所在列 |
| fieldDescColumn | int | 2 | 字段注释说明所在列 |
| charsetName | String | UTF-8 | 字符集编码 |
| genGetterAndSetter | boolean | true | 是否生成get和set方法 |
| genToString | boolean | true | 是否生成toString方法 |
| useLombok | boolean |false| 是否使用Lombok |
| parseEntityName | boolean |false| 是否解析实体名称 |
| entityNameRow | int | 0 | 实体名称所在行 |
| entityNameColumn | int | 0 | 实体名称所在列 |
| serializable | boolean | false | 是否实现Serializable序列化 |
| showHeader | boolean | true | 是否显示头注释 |
| header | String | Created by WordPOI | 头注释内容 |
| transformations | Map&lt;String,String&gt; |  | 需要转型的集合（自定义转型配置） |

### 代码示例

代码示例 (直接在main方法中调用即可)
```Java
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

```

### 文档示例

* 文档实体示例一（默认格式，见文档 Api1.docx）

1.1.	Result （响应结果实体）

| 字段    | 字段类型  |   说明  |
| :------| :------ | :------ |
|code    |  String | 0-代表成功，其它代表失败 |
|desc    |  String |	操作失败时的说明信息 |
|data    |	T	   |   返回对应的泛型<T>实体对象 |


1.2.	PageInfo （页码信息实体）

| 字段     |   字段类型    |   说明  |
| :------ | :------ | :------ |
| curPage   | Integer   |	当前页码 |
| pageSize  | Integer   |	页码大小，每一页的记录条数 |
| totalPage | Integer	| 总页数 | 
| hasNext   | Boolean	|  是否有下一页 | 
| data	    | List&lt;T&gt; | 	泛型T为对应的数据记录实体 | 


* 文档实体示例二（自动解析实体名称格式，见文档 Api3.docx）

1.1.    响应结果实体

<table>
    <tr>
    	<td colspan="3">Result</td>    
   </tr>
   <tr>
        <td>字段</td> 
        <td>字段类型</td> 
        <td>说明</td> 
   </tr>
   <tr>
        <td>code</td> 
        <td>String</td> 
        <td>0-代表成功，其它代表失败</td> 
   </tr>
   <tr>
        <td>desc</td> 
        <td>String</td> 
        <td>操作失败时的说明信息</td> 
   </tr>
   <tr>
        <td>data</td> 
        <td>T</td> 
        <td>返回对应的泛型&lt;T&gt;实体对象</td> 
   </tr>

</table>


1.2.    页码信息实体

<table>
    <tr>
    	<td colspan="3">PageInfo</td>    
   </tr>
   <tr>
        <td>字段</td> 
        <td>字段类型</td> 
        <td>说明</td> 
   </tr>
   <tr>
        <td>curPage</td> 
        <td>Integer</td> 
        <td>当前页码</td> 
   </tr>
   <tr>
        <td>curPage</td> 
        <td>Integer</td> 
        <td>当前页码</td> 
   </tr>
   <tr>
        <td>pageSize</td> 
        <td>Integer</td> 
        <td>页码大小，每一页的记录条数</td> 
   </tr>
   <tr>
        <td>totalPage</td> 
        <td>Integer</td> 
        <td>总页数</td> 
   </tr>
   <tr>
        <td>hasNext</td> 
        <td>Boolean</td> 
        <td>是否有下一页</td> 
   </tr>
   <tr>
        <td>data</td> 
        <td>List&lt;T&gt;</td> 
        <td>泛型T为对应的数据记录实体</td> 
   </tr>

</table>

更多使用详情，请查看[Test](src/test/java/Test.java)中的源码使用示例或直接查看[API帮助文档](https://jenly1314.github.io/projects/WordPOI/doc/)

## 版本记录

#### v1.0.1：2019-9-17
*  支持自动解析实体名称
*  支持添加自定义转型配置

#### v1.0.0：2019-6-12
*  WordPOI初始版本

## 赞赏
如果您喜欢WordPOI，或感觉WordPOI帮助到了您，可以点右上角“Star”支持一下，您的支持就是我的动力，谢谢 :smiley:
<p>您也可以扫描下面的二维码，请作者喝杯咖啡 :coffee:

<div>
   <img src="https://jenly1314.github.io/image/page/rewardcode.png">
</div>

## 关于我

| 我的博客                                                                                | GitHub                                                                                  | Gitee                                                                                  | CSDN                                                                                 | 博客园                                                                            |
|:------------------------------------------------------------------------------------|:----------------------------------------------------------------------------------------|:---------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------|
| <a title="我的博客" href="https://jenly1314.github.io" target="_blank">Jenly's Blog</a> | <a title="GitHub开源项目" href="https://github.com/jenly1314" target="_blank">jenly1314</a> | <a title="Gitee开源项目" href="https://gitee.com/jenly1314" target="_blank">jenly1314</a>  | <a title="CSDN博客" href="http://blog.csdn.net/jenly121" target="_blank">jenly121</a>  | <a title="博客园" href="https://www.cnblogs.com/jenly" target="_blank">jenly</a>  |

## 联系我

| 微信公众号        | Gmail邮箱                                                                          | QQ邮箱                                                                              | QQ群                                                                                                                       | QQ群                                                                                                                       |
|:-------------|:---------------------------------------------------------------------------------|:----------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------|
| [Jenly666](http://weixin.qq.com/r/wzpWTuPEQL4-ract92-R) | <a title="给我发邮件" href="mailto:jenly1314@gmail.com" target="_blank">jenly1314</a> | <a title="给我发邮件" href="mailto:jenly1314@vip.qq.com" target="_blank">jenly1314</a> | <a title="点击加入QQ群" href="https://qm.qq.com/cgi-bin/qm/qr?k=6_RukjAhwjAdDHEk2G7nph-o8fBFFzZz" target="_blank">20867961</a> | <a title="点击加入QQ群" href="https://qm.qq.com/cgi-bin/qm/qr?k=Z9pobM8bzAW7tM_8xC31W8IcbIl0A-zT" target="_blank">64020761</a> |

<div>
   <img src="https://jenly1314.github.io/image/page/footer.png">
</div>
