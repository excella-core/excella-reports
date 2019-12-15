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

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.model.ParsedReportInfo} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ParsedReportInfoTest {

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
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#getRowIndex()} のためのテスト・メソッド。
     */
    @Test
    public void testGetRowIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setRowIndex( 1);
        assertEquals( 1, info.getRowIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#setRowIndex(int)} のためのテスト・メソッド。
     */
    @Test
    public void testSetRowIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setRowIndex( 1);
        assertEquals( 1, info.getRowIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#getColumnIndex()} のためのテスト・メソッド。
     */
    @Test
    public void testGetColumnIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setColumnIndex( 1);
        assertEquals( 1, info.getColumnIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#setColumnIndex(int)} のためのテスト・メソッド。
     */
    @Test
    public void testSetColumnIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setColumnIndex( 1);
        assertEquals( 1, info.getColumnIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#getParsedObject()} のためのテスト・メソッド。
     */
    @Test
    public void testGetParsedObject() {
        ParsedReportInfo info = new ParsedReportInfo();
        
        Object object = new Object();
        info.setParsedObject(object);
        
        assertEquals( object, info.getParsedObject());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#setParsedObject(java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testSetParsedObject() {
        ParsedReportInfo info = new ParsedReportInfo();
        
        Object object = new Object();
        info.setParsedObject(object);
        
        assertEquals( object, info.getParsedObject());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#toString()} のためのテスト・メソッド。
     */
    @Test
    public void testToString() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setRowIndex( 1);
        info.setColumnIndex( 1);
        info.setDefaultRowIndex( 1);
        info.setDefaultColumnIndex( 1);
        Object object = new String("AAA");
        info.setParsedObject(object);
        
        assertEquals( "beforeLastCell=(1,1), afterLastCell=(1,1), AAA", info.toString());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#getDefaultRowIndex()} のためのテスト・メソッド。
     */
    @Test
    public void testGetDefaultRowIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setDefaultRowIndex( 1);
        assertEquals( 1, info.getDefaultRowIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#setDefaultRowIndex(int)} のためのテスト・メソッド。
     */
    @Test
    public void testSetDefaultRowIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setDefaultRowIndex( 1);
        assertEquals( 1, info.getDefaultRowIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#getDefaultColumnIndex()} のためのテスト・メソッド。
     */
    @Test
    public void testGetDefaultColumnIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setDefaultColumnIndex( 1);
        assertEquals( 1, info.getDefaultColumnIndex());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParsedReportInfo#setDefaultColumnIndex(int)} のためのテスト・メソッド。
     */
    @Test
    public void testSetDefaultColumnIndex() {
        ParsedReportInfo info = new ParsedReportInfo();
        info.setDefaultColumnIndex( 1);
        assertEquals( 1, info.getDefaultColumnIndex());
    }

}
