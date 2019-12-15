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

package org.bbreak.excella.reports.tag;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.fail;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.processor.ReportsCheckException;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.tag.SingleParamParser} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class SingleParamParserTest extends ReportsWorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version バージョン
     */
    public SingleParamParserTest( String version) {
        super( version);
    }

    /**
     * {@link org.bbreak.excella.reports.tag.SingleParamParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testParseSheetCellObject() {

        // -----------------------
        // □[正常系]通常解析
        // -----------------------

        Workbook workbook = getWorkbook();

        Sheet sheet1 = workbook.getSheetAt( 0);

        // テストデータ
        ParamInfo info = new ParamInfo();
        // 文字
        info.addParam( SingleParamParser.DEFAULT_TAG, "test", "test!");
        // 2バイト文字
        info.addParam( SingleParamParser.DEFAULT_TAG, "テスト", "テストです!");
        // byte
        info.addParam( SingleParamParser.DEFAULT_TAG, "test01", ( byte) 1);
        // Byte
        info.addParam( SingleParamParser.DEFAULT_TAG, "test02", new Byte( "1"));
        // short
        info.addParam( SingleParamParser.DEFAULT_TAG, "test03", ( short) 2);
        // Short
        info.addParam( SingleParamParser.DEFAULT_TAG, "test04", new Short( "2"));
        // int
        info.addParam( SingleParamParser.DEFAULT_TAG, "test05", 3);
        // Integer
        info.addParam( SingleParamParser.DEFAULT_TAG, "test06", new Integer( "3"));
        // long
        info.addParam( SingleParamParser.DEFAULT_TAG, "test07", 4L);
        // Long
        info.addParam( SingleParamParser.DEFAULT_TAG, "test08", new Long( "4"));
        // float
        info.addParam( SingleParamParser.DEFAULT_TAG, "test09", 5.1f);
        // Float
        info.addParam( SingleParamParser.DEFAULT_TAG, "test10", new Float( "5.1"));
        // double
        info.addParam( SingleParamParser.DEFAULT_TAG, "test11", 6.1);
        // Double
        info.addParam( SingleParamParser.DEFAULT_TAG, "test12", new Double( 6.1));
        // BigInteger
        info.addParam( SingleParamParser.DEFAULT_TAG, "test13", new BigInteger( "7"));
        // BigDecimal
        info.addParam( SingleParamParser.DEFAULT_TAG, "test14", new BigDecimal( "8.5"));
        // Date
        Calendar cal = Calendar.getInstance();
        cal.clear();
        cal.set( 2009, 4, 22);
        info.addParam( SingleParamParser.DEFAULT_TAG, "test15", cal.getTime());

        SingleParamParser parser = new SingleParamParser();

        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();

        reportsParserInfo.setParamInfo( info);

        // 解析処理
        try {
            parseSheet( parser, sheet1, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook();
        Sheet expectedSheet = expectedWorkbook.getSheet( "Sheet1");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet1, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        }

        // -----------------------
        // □[正常系]タグ名の変更
        // -----------------------

        workbook = getWorkbook();

        // テストデータ
        ParamInfo infoP = new ParamInfo();
        infoP.addParam( "$P", "test", "test!");
        infoP.addParam( "$P", "テスト", "テストです!");
        infoP.addParam( "$P", "test01", ( byte) 1);
        infoP.addParam( "$P", "test02", new Byte( "1"));
        infoP.addParam( "$P", "test03", ( short) 2);
        infoP.addParam( "$P", "test04", new Short( "2"));
        infoP.addParam( "$P", "test05", 3);
        infoP.addParam( "$P", "test06", new Integer( "3"));
        infoP.addParam( "$P", "test07", 4L);
        infoP.addParam( "$P", "test08", new Long( "4"));
        infoP.addParam( "$P", "test09", 5.1f);
        infoP.addParam( "$P", "test10", new Float( "5.1"));
        infoP.addParam( "$P", "test11", 6.1);
        infoP.addParam( "$P", "test12", new Double( 6.1));
        infoP.addParam( "$P", "test13", new BigInteger( "7"));
        infoP.addParam( "$P", "test14", new BigDecimal( "8.5"));
        cal.clear();
        cal.set( 2009, 4, 22);
        infoP.addParam( "$P", "test15", cal.getTime());

        // コンストラクタによるタグ名セット
        parser = new SingleParamParser( "$P");

        reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( infoP);

        Sheet sheet2 = workbook.getSheetAt( 1);

        // 解析処理
        try {
            parseSheet( parser, sheet2, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.getSheet( "Sheet2");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet2, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        }

        // セッターによるタグ名セット
        parser = new SingleParamParser();
        parser.setTag( "$P");

        sheet2 = workbook.getSheetAt( 1);

        // 解析処理
        try {
            parseSheet( parser, sheet2/** sheet1 */
            , reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.getSheet( "Sheet2");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet2, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        }

        // -----------------------
        // □[正常系]必須パラメータがない場合
        // -----------------------
        parser = new SingleParamParser();

        Sheet sheet3 = workbook.getSheetAt( 2);

        // 解析処理
        try {
            parseSheet( parser, sheet3, reportsParserInfo);
        } catch ( ParseException e) {

        }
        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.getSheet( "Sheet3");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet3, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        }
    }

    /**
     * {@link org.bbreak.excella.reports.tag.SingleParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        // -----------------------
        // □制御行使用有無：無
        // -----------------------
        SingleParamParser parser = new SingleParamParser();
        assertFalse( parser.useControlRow());
    }

}
