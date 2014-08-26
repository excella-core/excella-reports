/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: RemoveAdapterTest.java 84 2009-10-30 04:07:19Z akira-yokoi $
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
package org.bbreak.excella.reports.listener;

import static org.junit.Assert.*;

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
     * {@link org.bbreak.excella.reports.listener.RemoveAdapter#postParse(org.apache.poi.ss.usermodel.Sheet, org.bbreak.excella.core.SheetParser, org.bbreak.excella.core.SheetData)} のためのテスト・メソッド。
     */
    @Test
    public void testPostParse() {

        Workbook workbook = getWorkbook();

        RemoveAdapter adapter = new RemoveAdapter();

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

        workbook = getWorkbook();
        sheet = workbook.getSheetAt( 1);
        try {
            adapter.postParse( sheet, sheetParser, null);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail();
        }
        checkSheet( workbook.getSheetName( 1), sheet, true);

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
