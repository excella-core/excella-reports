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

package org.bbreak.excella.reports.listener;

import static org.junit.Assert.fail;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.processor.ReportCreateHelper;
import org.bbreak.excella.reports.processor.ReportsCheckException;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.listener.BreakAdapter} のためのテスト・クラス。
 * 
 * @author T.Maruyama
 */
public class BreakAdapterTest extends ReportsWorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version
     */
    public BreakAdapterTest( String version) {
        super( version);
    }

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
     * {@link org.bbreak.excella.reports.listener.BreakAdapter#postParse(org.apache.poi.ss.usermodel.Sheet, org.bbreak.excella.core.SheetParser, org.bbreak.excella.core.SheetData)} のためのテスト・メソッド。
     */
    @Test
    public void testPostParse() {

        Workbook workbook = getWorkbook();

        BreakAdapter adapter = new BreakAdapter();

        SheetParser sheetParser = new SheetParser();
        List<ReportsTagParser<?>> reportParsers = new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values());

        for ( ReportsTagParser<?> parser : reportParsers) {
            sheetParser.addTagParser( parser);
        }

        Sheet sheet = workbook.getSheetAt( 0);
        try {
            adapter.postParse( sheet, sheetParser, null);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail();
        }
        checkSheet( workbook.getSheetName( 0), sheet, true);
    }

    private void checkSheet( String expectedSheetName, Sheet actualSheet, boolean outputExcel) {

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook();
        Sheet expectedSheet = expectedWorkbook.getSheet( expectedSheetName);

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, actualSheet, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            String tmpDirPath = ReportsTestUtil.getTestOutputDir();
            try {
                String filepath = null;
                Date now = new Date();
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + this.getClass().getSimpleName() + now.getTime() + ".xlsx";
                } else {
                    filepath = tmpDirPath + this.getClass().getSimpleName() + now.getTime() + ".xls";
                }
                PoiUtil.writeBook( actualSheet.getWorkbook(), filepath);

            } catch ( IOException e) {
                e.printStackTrace();
            }
        }
    }
}
