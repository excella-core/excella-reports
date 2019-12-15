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

package org.bbreak.excella.reports.model;

import static org.junit.Assert.*;

import java.util.ArrayList;
import java.util.List;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.model.ReportBook} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ReportBookTest {

    private ReportBook reportBook = null;
    
    /**
     * @throws java.lang.Exception
     */
    @BeforeClass
    public static void setUpBeforeClass() throws Exception {
    }

    /**
     * @throws java.lang.Exception
     */
    @Before
    public void setUp() throws Exception {
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#ReportBook(String, String, ConvertConfiguration...)} のためのテスト・メソッド。
     */
    @Test
    public void testReportBookStringStringConvertConfiguration() {
        
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);
        
        assertEquals( "c:/test", reportBook.getOutputFileName());
        assertArrayEquals( configurations, reportBook.getConfigurations());
    }
    
    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#ReportBook(String, String, String...)} のためのテスト・メソッド。
     */
    @Test
    public void testReportBookStringStringString() {
        
         reportBook = new ReportBook("", "c:/test", "AA");
         
        
        assertEquals( "c:/test", reportBook.getOutputFileName());
        
        
        assertEquals( "AA", reportBook.getConfigurations()[0].getFormatType());
    }


    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#getOutputFileName()} のためのテスト・メソッド。
     */
    @Test
    public void testGetOutputFileName() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);

        assertEquals( "c:/test", reportBook.getOutputFileName());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#getConfigurations()} のためのテスト・メソッド。
     */
    @Test
    public void testGetConfigurations() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);
        
        assertArrayEquals( configurations, reportBook.getConfigurations());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#addReportSheet(org.bbreak.excella.reports.model.ReportSheet)} のためのテスト・メソッド。
     */
    @Test
    public void testAddReportSheet() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);

        ReportSheet reportSheet = new ReportSheet("");
        reportBook.addReportSheet( reportSheet);
        assertTrue( reportBook.getReportSheets().contains( reportSheet));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#addReportSheets(java.util.List)} のためのテスト・メソッド。
     */
    @Test
    public void testAddReportSheets() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);

        List<ReportSheet> list = new ArrayList<ReportSheet>();
        
        ReportSheet reportSheet1 = new ReportSheet("");
        list.add( reportSheet1);
        ReportSheet reportSheet2 = new ReportSheet("");
        list.add( reportSheet2);
        ReportSheet reportSheet3 = new ReportSheet("");
        list.add( reportSheet3);
        
        
        reportBook.addReportSheets( list);
        assertTrue( reportBook.getReportSheets().containsAll( list));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#getReportSheets()} のためのテスト・メソッド。
     */
    @Test
    public void testGetReportSheets() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);

        List<ReportSheet> list = new ArrayList<ReportSheet>();
        
        ReportSheet reportSheet1 = new ReportSheet("");
        list.add( reportSheet1);
        ReportSheet reportSheet2 = new ReportSheet("");
        list.add( reportSheet2);
        ReportSheet reportSheet3 = new ReportSheet("");
        list.add( reportSheet3);
        
        reportBook.addReportSheets( list);
        
        assertArrayEquals( list.toArray(), reportBook.getReportSheets().toArray());
    }


    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#removeReportSheet(org.bbreak.excella.reports.model.ReportSheet)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveReportSheet() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);

        List<ReportSheet> list = new ArrayList<ReportSheet>();
        
        ReportSheet reportSheet1 = new ReportSheet("");
        list.add( reportSheet1);
        ReportSheet reportSheet2 = new ReportSheet("");
        list.add( reportSheet2);
        ReportSheet reportSheet3 = new ReportSheet("");
        list.add( reportSheet3);
        
        reportBook.addReportSheets( list);
        
        
        
        reportBook.removeReportSheet( reportSheet1);
        list.remove( reportSheet1);
        
        assertArrayEquals( list.toArray(), reportBook.getReportSheets().toArray());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportBook#clearReportSheets()} のためのテスト・メソッド。
     */
    @Test
    public void testClearReportSheets() {
        ConvertConfiguration[] configurations = new ConvertConfiguration[]{new ConvertConfiguration("")};
        reportBook = new ReportBook("", "c:/test", configurations);
        
        List<ReportSheet> list = new ArrayList<ReportSheet>();
        
        ReportSheet reportSheet1 = new ReportSheet("");
        list.add( reportSheet1);
        ReportSheet reportSheet2 = new ReportSheet("");
        list.add( reportSheet2);
        ReportSheet reportSheet3 = new ReportSheet("");
        list.add( reportSheet3);
        
        reportBook.addReportSheets( list);


        reportBook.clearReportSheets();
        assertTrue( reportBook.getReportSheets().isEmpty());
    }



}
