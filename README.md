# excel-parser
根据当前目录config.xml目录，抽取excel表格数据

# 使用方法

```
mvn clean install -DskipTests
```

构建项目，生成可依赖jar，在jar目录下新建config.xml 格式参照/src/resources/config.xml，新建excel目录，把需要抽取的excel文件放到excel文件夹下

# 运行

```
java -jar excel-parser.jar
```
会在本地目录生成output.xlsx
