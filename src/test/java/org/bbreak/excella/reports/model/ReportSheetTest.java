/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportSheetTest.java 84 2009-10-30 04:07:19Z akira-yokoi $
 * $Revision: 84 $
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

import java.util.HashMap;
import java.util.Map;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.model.ReportSheet} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ReportSheetTest {

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
     * {@link org.bbreak.excella.reports.model.ReportSheet#ReportSheet(java.lang.String, java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testReportSheetStringString() {
        ReportSheet reportSheet = new ReportSheet( "テンプレートシート", "出力シート");
        assertEquals( "テンプレートシート", reportSheet.getTemplateName());
        assertEquals( "出力シート", reportSheet.getSheetName());

    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#ReportSheet(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testReportSheetString() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        assertEquals( "シート", reportSheet.getTemplateName());
        assertEquals( "シート", reportSheet.getSheetName());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#getTemplateName()} のためのテスト・メソッド。
     */
    @Test
    public void testGetTemplateName() {
        ReportSheet reportSheet = new ReportSheet( "テンプレートシート", "出力シート");
        assertEquals( "テンプレートシート", reportSheet.getTemplateName());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#getSheetName()} のためのテスト・メソッド。
     */
    @Test
    public void testGetSheetName() {
        ReportSheet reportSheet = new ReportSheet( "テンプレートシート", "出力シート");
        assertEquals( "出力シート", reportSheet.getSheetName());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#setSheetName(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testSetSheetName() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        reportSheet.setSheetName( "出力シート");
        assertEquals( "出力シート", reportSheet.getSheetName());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#setTemplateName(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testSetTemplateName() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        reportSheet.setTemplateName( "テンプレートシート");
        assertEquals( "テンプレートシート", reportSheet.getTemplateName());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#getParamInfo()} のためのテスト・メソッド。
     */
    @Test
    public void testGetParamInfo() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "test1", "データ1");
        map.put( "test2", "データ2");
        map.put( "test3", "データ3");
        reportSheet.addParams( "$", map);
        reportSheet.addParam( "$A", "testA", "データA");

        ParamInfo info = reportSheet.getParamInfo();

        assertEquals( "データ1", info.getParam( "$", "test1"));
        assertEquals( "データ2", info.getParam( "$", "test2"));
        assertEquals( "データ3", info.getParam( "$", "test3"));
        assertEquals( "データA", info.getParam( "$A", "testA"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#addParam(java.lang.String, java.lang.String, java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testAddParam() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        reportSheet.addParam( "$", "test", "データ");

        assertEquals( "データ", reportSheet.getParam( "$", "test"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#addParams(java.lang.String, java.util.Map)} のためのテスト・メソッド。
     */
    @Test
    public void testAddParams() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "test1", "データ1");
        map.put( "test2", "データ2");
        map.put( "test3", "データ3");
        reportSheet.addParams( "$", map);

        assertEquals( "データ1", reportSheet.getParam( "$", "test1"));
        assertEquals( "データ2", reportSheet.getParam( "$", "test2"));
        assertEquals( "データ3", reportSheet.getParam( "$", "test3"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#clearParam()} のためのテスト・メソッド。
     */
    @Test
    public void testClearParam() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "test1", "データ1");
        map.put( "test2", "データ2");
        map.put( "test3", "データ3");
        reportSheet.addParams( "$", map);
        reportSheet.addParam( "$A", "testA", "データA");

        reportSheet.clearParam();

        assertNull( reportSheet.getParam( "$", "test1"));
        assertNull( reportSheet.getParam( "$", "test2"));
        assertNull( reportSheet.getParam( "$", "test3"));
        assertNull( reportSheet.getParam( "$A", "testA"));

    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#clearParam(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testClearParamString() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "test1", "データ1");
        map.put( "test2", "データ2");
        map.put( "test3", "データ3");
        reportSheet.addParams( "$", map);
        reportSheet.addParam( "$A", "testA", "データA");

        reportSheet.clearParam( "$");
        assertNull( reportSheet.getParam( "$", "test1"));
        assertNull( reportSheet.getParam( "$", "test2"));
        assertNull( reportSheet.getParam( "$", "test3"));

        assertEquals( "データA", reportSheet.getParam( "$A", "testA"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#getParam(java.lang.String, java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testGetParam() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "test1", "データ1");
        map.put( "test2", "データ2");
        map.put( "test3", "データ3");
        reportSheet.addParams( "$", map);
        reportSheet.addParam( "$A", "testA", "データA");

        assertEquals( "データ1", reportSheet.getParam( "$", "test1"));
        assertEquals( "データ2", reportSheet.getParam( "$", "test2"));
        assertEquals( "データ3", reportSheet.getParam( "$", "test3"));
        assertEquals( "データA", reportSheet.getParam( "$A", "testA"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ReportSheet#removeParam(java.lang.String, java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveParam() {
        ReportSheet reportSheet = new ReportSheet( "シート");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "test1", "データ1");
        map.put( "test2", "データ2");
        map.put( "test3", "データ3");
        reportSheet.addParams( "$", map);
        reportSheet.addParam( "$A", "testA", "データA");

        reportSheet.removeParam( "$", "test2");
        reportSheet.removeParam( "$A", "testA");

        assertEquals( "データ1", reportSheet.getParam( "$", "test1"));
        assertNull( reportSheet.getParam( "$", "test2"));
        assertEquals( "データ3", reportSheet.getParam( "$", "test3"));
        assertNull( reportSheet.getParam( "$A", "testA"));
    }

}
