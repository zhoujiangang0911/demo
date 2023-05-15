/*
 * Copyright 2013-2018 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.example.demo.demos.web;

import com.alibaba.fastjson.JSON;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.*;

/**
 * @author <a href="mailto:chenxilzx1@gmail.com">theonefx</a>
 */
@Controller
public class BasicController {

    private static final Logger LOG = LoggerFactory.getLogger(BasicController.class);

    // http://127.0.0.1:8080/hello?name=lisi
    @RequestMapping("/hello")
    @ResponseBody
    public String hello(@RequestParam(name = "name", defaultValue = "unknown user") String name) {
        return "Hello " + name;
    }

    // http://127.0.0.1:8080/user
    @RequestMapping("/user")
    @ResponseBody
    public User user(HttpServletRequest request, HttpServletResponse response) {
        LOG.debug("debug");
        LOG.info("info");
        LOG.warn("warn");
        User user = new User();
        user.setName("theonefx");
        user.setAge(666);
        String xlsType = "2";
        XSSFWorkbook workbook = null;
        Workbook sBook = null;
        Sheet sheet = null;
        BufferedOutputStream out = null;
        try {
            String fileName = request.getParameter("genFileName");
            if (fileName == null || fileName.length() == 0) {
                fileName = "test";
            }
            response.setContentType("application/vnd.ms-excel");//application/force-download
            if ("1".equals(xlsType)) {
                //Office 2003EXCEL
                response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(fileName + ".xls", "UTF8"));
            } else if ("2".equals(xlsType)) {
                //Office 2007EXCEL
                response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(fileName + ".xlsx", "UTF8"));
            }

            out = new BufferedOutputStream(response.getOutputStream());

            if ("1".equals(xlsType)) {
                LOG.info("=========开始创建EXCEL 2003的文件   begin=========");
                sBook = new HSSFWorkbook();
                sheet = sBook.createSheet("sheet1");
                LOG.info("=========开始创建EXCEL 2003的文件   end=========");
            } else if ("2".equals(xlsType)) {
                LOG.info("=========开始创建EXCEL 2007的文件   begin=========");
                sBook = new SXSSFWorkbook(100);
                sheet = sBook.createSheet("sheet1");
                LOG.info("=========开始创建EXCEL 2007的文件   end=========");
            }


            LOG.debug("导出xls开始...");
            LinkedHashMap<String, String> columnMap = new LinkedHashMap<String, String>();




            sheet.setForceFormulaRecalculation(true);
            // 写表头
            int headerCellIndex = 0;
            Row headerRow = sheet.createRow(0);
            if (!columnMap.isEmpty()) {
                CellStyle indexStyle = null;
                for (Iterator iterator = columnMap.keySet().iterator(); iterator.hasNext(); ) {
                    String key = iterator.next().toString();// 英文
                    Cell headerCell = headerRow.createCell(headerCellIndex);
                    genCellValue(headerCell, columnMap.get(key));
                    indexStyle = sheet.getColumnStyle(headerCellIndex);
                    if (indexStyle != null) {
                        headerCell.setCellStyle(indexStyle);
                    }
                    headerCellIndex++;
                }
            }
            int row = 0;// excel行数
            Row bodyRow = null;


                /** 获取配置中的导出总条数，进行设定     ----begin----**/
                int exportTotal = 0;

                /** 获取配置中的导出总条数，进行设定     ----end----**/


                sBook.write(out);
                out.flush();
//				sBook.close();
        } catch (Exception e) {
            LOG.error(e.getMessage(), e);
            if (sBook != null && out != null) {
                try {
                    sBook.write(out);
                    out.flush();
//					sBook.close();
                } catch (IOException e1) {
                    LOG.error(e1.getMessage(), e1);
                }
            }
        } finally {
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                LOG.error(e.getMessage(), e);
            }
        }
        return user;
    }

    /**
     * 遍历对象属性与表头匹配
     *
     * @param targetObj     需要插入的对象
     * @param gridTitleList 表头List
     * @param Row           row行
     */
    private void writeSheet(Object targetObj, List<String> gridTitleList, Row row, Sheet sheet) throws Exception {
        Map<Object, Object> fieldMap = JSON.parseObject(JSON.toJSONString(targetObj), Map.class);
        Cell cell = null;
        CellStyle style = null;
        for (Map.Entry entry : fieldMap.entrySet()) {// 循环属性匹配表头
            String fieldKey = entry.getKey().toString();
            int index = gridTitleList.indexOf(fieldKey);
            if (index != -1) {
                cell = row.createCell(index);
                genCellValue(cell, entry.getValue());
                style = sheet.getColumnStyle(index);
                if (style != null) {
                    cell.setCellStyle(style);
                }
            }
        }
    }
    /**
     * 为单元格生成cell值
     *
     * @param cell
     * @param value
     */
    private void genCellValue(Cell cell, Object value) {
        if (value instanceof String) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(value.toString());
        } else if (value instanceof Number) {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
            return;
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
            return;
        } else {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(value.toString());
        }
    }


    // http://127.0.0.1:8080/save_user?name=newName&age=11
    @RequestMapping("/save_user")
    @ResponseBody
    public String saveUser(User u) {
        return "user will save: name=" + u.getName() + ", age=" + u.getAge();
    }

    // http://127.0.0.1:8080/html
    @RequestMapping("/html")
    public String html(){
        return "index.html";
    }

    @ModelAttribute
    public void parseUser(@RequestParam(name = "name", defaultValue = "unknown user") String name
            , @RequestParam(name = "age", defaultValue = "12") Integer age, User user) {
        user.setName("zhangsan");
        user.setAge(18);
    }
}
