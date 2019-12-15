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
