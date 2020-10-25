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
 * {@link org.bbreak.excella.reports.listener.RemoveAdapter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class RemoveAdapterTest extends ReportsWorkbookTest {

    public RemoveAdapterTest( String version) {
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
     * {@link org.bbreak.excella.reports.listener.RemoveAdapter#postParse(org.apache.poi.ss.usermodel.Sheet, org.bbreak.excella.core.SheetParser, org.bbreak.excella.core.SheetData)}
     * のためのテスト・メソッド。
     * 
     * @throws ParseException
     * @throws ReportsCheckException
     */
    @Test
    public void testPostParse() throws ParseException, ReportsCheckException {

        Workbook workbook = getWorkbook();

        RemoveAdapter adapter = new RemoveAdapter();

        SheetParser sheetParser = new SheetParser();
        List<ReportsTagParser<?>> reportParsers = new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values());

        for ( ReportsTagParser<?> parser : reportParsers) {
            sheetParser.addTagParser( parser);
        }

        Sheet sheet = workbook.getSheetAt( 0);
        adapter.postParse( sheet, sheetParser, null);
        checkSheet( workbook.getSheetName( 0), sheet, true);

        workbook = getWorkbook();
        sheet = workbook.getSheetAt( 1);
        adapter.postParse( sheet, sheetParser, null);
        checkSheet( workbook.getSheetName( 1), sheet, true);

    }

    private void checkSheet( String expectedSheetName, Sheet actualSheet, boolean outputExcel)
            throws ReportsCheckException {

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook();
        Sheet expectedSheet = expectedWorkbook.getSheet( expectedSheetName);

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, actualSheet, false);
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
