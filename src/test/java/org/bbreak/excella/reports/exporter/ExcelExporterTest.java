/*-
 * #%L
 * excella-reports
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.reports.exporter;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.ExcelExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 *
 */
public class ExcelExporterTest {
    
    private Workbook hssfWb = new HSSFWorkbook();

    private Workbook xssfWb = new XSSFWorkbook();
    
    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();
    
    ConvertConfiguration configuration = null;
    
    
    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)} のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {
        
        ExcelExporter exporter = new ExcelExporter();
        String filePath = null;
        
        //XLS
        configuration = new ConvertConfiguration(ExcelExporter.EXTENTION_XLS);
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        try {
            exporter.output( hssfWb, new BookData(), configuration);
            File file = new File(exporter.getFilePath());
            assertTrue( file.exists());
        } catch ( ExportException e) {
            e.printStackTrace();
        }
        
        
        //XLSX
        configuration = new ConvertConfiguration(ExcelExporter.EXTENTION_XLSX);
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        try {
            exporter.output( xssfWb, new BookData(), configuration);
            File file = new File(exporter.getFilePath());
            assertTrue( file.exists());
            
        } catch ( ExportException e) {
            e.printStackTrace();
        }
        
        
        
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        ExcelExporter exporter = new ExcelExporter();
        assertEquals( "EXCEL", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        ExcelExporter exporter = new ExcelExporter();
        assertEquals( "", exporter.getExtention());
    }
    
}
