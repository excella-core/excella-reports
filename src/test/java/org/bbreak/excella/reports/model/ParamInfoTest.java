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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.HashMap;
import java.util.Map;

import org.junit.Before;
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
