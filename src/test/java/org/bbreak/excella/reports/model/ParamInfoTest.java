/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ParamInfoTest.java 201 2013-03-15 05:11:47Z kamisono_bb $
 * $Revision: 201 $
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
 * {@link org.bbreak.excella.reports.model.ParamInfo} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ParamInfoTest {
    
    private static ParamInfo info = null;
    
    /**
     * @throws java.lang.Exception
     */
    @Before
    public void setUp() throws Exception {
        info = new ParamInfo();
        info.addParam( "$", "Test0", "Data0");
        info.addParam( "$", "Test1", "Data1");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("Test2", "Data2");
        map.put("Test3", "Data3");
        map.put("Test4", "Data4");
        info.addParams("$Map", map);
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParamInfo#getParam(java.lang.String, java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testGetParam() {
        
        assertEquals( "Data0", info.getParam( "$", "Test0"));
        
    }
    /**
     * {@link org.bbreak.excella.reports.model.ParamInfo#addParam(java.lang.String, java.lang.String, java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testAddParam() {
        assertEquals( "Data1", info.getParam( "$", "Test1"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParamInfo#addParams(java.lang.String, java.util.Map)} のためのテスト・メソッド。
     */
    @Test
    public void testAddParams() {
        assertEquals( "Data2", info.getParam( "$Map", "Test2"));
        assertEquals( "Data3", info.getParam( "$Map", "Test3"));
        assertEquals( "Data4", info.getParam( "$Map", "Test4"));
    }



    /**
     * {@link org.bbreak.excella.reports.model.ParamInfo#removeParam(java.lang.String, java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveParam() {
        
        info.removeParam( "$Map", "Test4");
        assertNull( info.getParam( "$Map", "Test4"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParamInfo#clearParam(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testClearParamString() {
        info.clearParam( "$Map");
        assertNull( info.getParam( "$Map", "Test2"));
        assertNull( info.getParam( "$Map", "Test3"));
        assertEquals( "Data1", info.getParam( "$", "Test1"));
    }

    /**
     * {@link org.bbreak.excella.reports.model.ParamInfo#clearParam()} のためのテスト・メソッド。
     */
    @Test
    public void testClearParam() {
        info.clearParam();
        assertNull( info.getParam( "$", "Test0"));
        assertNull( info.getParam( "$", "Test1"));

    }

}
