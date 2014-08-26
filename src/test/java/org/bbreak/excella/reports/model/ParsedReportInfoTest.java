/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ParsedReportInfoTest.java 34 2009-09-29 03:25:22Z tsuchida $
 * $Revision: 34 $
 *
 * This file is part of ExCella Reports.
 *
 * ExCella Reports is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License version 3
 * only, as published by the Free Software Foundation.
 *
 * ExCella Reports is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License version 3 for more details
 * (a copy is included in the COPYING.LESSER file that accompanied this code).
 *
 * You should have received a copy of the GNU Lesser General Public License
 * version 3 along with ExCella Reports.  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
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
