# demo
解决weblogic部署poi问题
部署app.war
访问app/user
下载test.xlsx能打开说明成功了

----
weblog.xml添加
weblogic 启动添加  
-Djava.awt.headless=true
```
<prefer-application-packages>
    <package-name>org.apache.commons.collections4.*</package-name>
    <package-name>org.apache.commons.compress.*</package-name>
    <package-name>org.apache.poi.*</package-name>
    <package-name>org.apache.xmlbeans.*</package-name>
    <package-name>org.openxmlformats.*</package-name>
    <package-name>schemaorg_apache_xmlbeans.*</package-name>
</prefer-application-packages>
<prefer-application-resources>
    <resource-name>schemaorg_apache_xmlbeans/system/sXMLCONFIG/TypeSystemHolder.class</resource-name>
    <resource-name>schemaorg_apache_xmlbeans/system/sXMLLANG/TypeSystemHolder.class</resource-name>
    <resource-name>schemaorg_apache_xmlbeans/system/sXMLSCHEMA/TypeSystemHolder.class</resource-name>
    <resource-name>schemaorg_apache_xmlbeans/system/sXMLTOOLS/TypeSystemHolder.class</resource-name>
</prefer-application-resources>
```
成功解决问题

