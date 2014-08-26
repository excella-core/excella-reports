/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ConvertConfigurationTest.java 84 2009-10-30 04:07:19Z akira-yokoi $
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
 * {@link org.bbreak.excella.reports.model.ConvertConfiguration} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ConvertConfigurationTest {

    private ConvertConfiguration configuration = null;

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
        configuration = new ConvertConfiguration( "Test");
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#ConvertConfiguration(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testConvertConfiguration() {

        assertEquals( "Test", configuration.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        assertEquals( "Test", configuration.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#setFormatType(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testSetFormatType() {
        configuration.setFormatType( "Test1");
        assertEquals( "Test1", configuration.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#addOption(java.lang.String, java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testAddOption() {
        configuration.addOption( "option1", true);
        assertEquals( true, configuration.getOptionValue( "option1"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#addOptions(java.util.Map)} のためのテスト・メソッド。
     */
    @Test
    public void testAddOptions() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "option2", 1);
        map.put( "option3", "test");
        map.put( "option4", "テスト");
        configuration.addOptions( map);

        assertEquals( 1, configuration.getOptionValue( "option2"));
        assertEquals( "test", configuration.getOptionValue( "option3"));
        assertEquals( "テスト", configuration.getOptionValue( "option4"));

    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#removeOption(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveOption() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "option2", 1);
        map.put( "option3", "test");
        map.put( "option4", "テスト");
        configuration.addOptions( map);

        configuration.removeOption( "option2");
        assertNull( configuration.getOptionValue( "option2"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#clearOptions()} のためのテスト・メソッド。
     */
    @Test
    public void testClearOptions() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "option2", 1);
        map.put( "option3", "test");
        map.put( "option4", "テスト");
        configuration.addOptions( map);

        configuration.clearOptions();
        assertTrue( configuration.getOptions().isEmpty());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#getOptionsProperties()} のためのテスト・メソッド。
     */
    @Test
    public void testGetOptionsProperties() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "option2", 1);
        map.put( "option3", "test");
        map.put( "option4", "テスト");
        configuration.addOptions( map);

        assertArrayEquals( map.keySet().toArray(), configuration.getOptionsProperties().toArray());
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#getOptionValue(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testGetOptionValue() {
        configuration.addOption( "option1", true);
        assertEquals( true, configuration.getOptionValue( "option1"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ConvertConfiguration#getOptions()} のためのテスト・メソッド。
     */
    @Test
    public void testGetOptions() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put( "option2", "1");
        map.put( "option3", "test");
        map.put( "option4", "テスト");
        configuration.addOptions( map);

        assertArrayEquals( map.keySet().toArray(), configuration.getOptions().keySet().toArray());
        assertArrayEquals( map.values().toArray(), configuration.getOptions().values().toArray());
    }

}
