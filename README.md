# demo
解决weblogic部署poi问题
部署app.war
访问app/user
下载test.xlsx能打开说明成功了

----
weblog.xml添加
```
 <wls:prefer-application-resources>
            <wls:resource-name>schemaorg_apache_xmlbeans/*</wls:resource-name>
        </wls:prefer-application-resources>
```
成功解决问题
