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

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Collections;
import java.util.Set;
import java.util.TreeSet;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.processor.ReportsCheckException;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.Assert;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.tag.ImageParamParser} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ImageParamParserTest extends ReportsWorkbookTest {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( ImageParamParser.class);

    /**
     * テンプレートシート(sheet5)のシートindex
     */
    private static final int INDEX_OF_SHEET5 = 3;

    /**
     * テンプレートシート(sheet6)のシートindex
     */
    private static final int INDEX_OF_SHEET6 = 4;

    /**
     * テンプレートシート(sheet5)のシートコピー数
     */
    private static final int PLURAL_COPY_FIRST_NUM_OF_SHEETS = 2;

    /**
     * テンプレートシート(sheet6)のシートコピー数
     */
    private static final int PLURAL_COPY_SECOND_NUM_OF_SHEETS = 2;

    /**
     * コンストラクタ
     * 
     * @param version バージョン
     */
    public ImageParamParserTest( String version) {
        super( version);
    }

    /**
     * {@link org.bbreak.excella.reports.tag.ImageParamParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)} のためのテスト・メソッド。
     * 
     * @throws IOException
     * @throws EncryptedDocumentException
     */
    @Test
    public void testParseSheetCellObject() throws EncryptedDocumentException, IOException {
        // -----------------------
        // □[正常系]通常解析
        // -----------------------
        Workbook workbook = getWorkbook();

        Sheet sheet1 = workbook.getSheetAt( 0);

        ImageParamParser parser = new ImageParamParser();

        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( createTestData( ImageParamParser.DEFAULT_TAG));

        try {
            parseSheet( parser, sheet1, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 不要シートの削除
        Set<Integer> delSheetIndexs = new TreeSet<Integer>( Collections.reverseOrder());
        for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if ( i != 0) {
                delSheetIndexs.add( i);
            }
        }
        for ( Integer i : delSheetIndexs) {
            workbook.removeSheetAt( i);
        }

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook();
        Sheet expectedSheet = expectedWorkbook.getSheet( "Sheet1");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet1, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());

        } finally {
            // オブジェクトはチェックできないので、出力して確認
            String tmpDirPath = System.getProperty( "user.dir") + "/work/test/"; // System.getProperty( "java.io.tmpdir");
            File file = new File( tmpDirPath);
            if ( !file.exists()) {
                file.mkdirs();
            }

            try {
                String filepath = null;
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + "ImageParamParserTest1.xlsx";
                } else {
                    filepath = tmpDirPath + "ImageParamParserTest1.xls";
                }
                PoiUtil.writeBook( workbook, filepath);
                log.info( "出力ファイルを確認してください：" + filepath);

            } catch ( IOException e) {
                fail( e.toString());
            }
        }

        // -----------------------
        // □[正常系]タグ変更
        // -----------------------
        workbook = getWorkbook();

        Sheet sheet2 = workbook.getSheetAt( 1);

        parser = new ImageParamParser( "$Image");

        reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( createTestData( "$Image"));

        try {
            parseSheet( parser, sheet2, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 不要シートの削除
        delSheetIndexs.clear();
        for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if ( i != 1) {
                delSheetIndexs.add( i);
            }
        }
        for ( Integer i : delSheetIndexs) {
            workbook.removeSheetAt( i);
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.getSheet( "Sheet2");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet2, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            // オブジェクトはチェックできないので、出力して確認
            String tmpDirPath = System.getProperty( "user.dir") + "/work/test/"; // System.getProperty( "java.io.tmpdir");;
            File file = new File( tmpDirPath);
            if ( !file.exists()) {
                file.mkdir();
            }
            try {
                String filepath = null;
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + "ImageParamParserTest2.xlsx";
                } else {
                    filepath = tmpDirPath + "ImageParamParserTest2.xls";
                }
                PoiUtil.writeBook( workbook, filepath);
                log.info( "出力ファイルを確認してください：" + filepath);

            } catch ( IOException e) {
                fail( e.toString());
            }
        }

        String filename = this.getClass().getSimpleName();
        if ( version.equals( "2007")) {
            filename = filename + "_err.xlsx";
        } else if ( version.equals( "2003")) {
            filename = filename + "_err.xls";
        }

        URL url = this.getClass().getResource( filename);
        String path = null;
        try {
            path = URLDecoder.decode( url.getFile(), "UTF-8");

        } catch ( UnsupportedEncodingException e) {
            Assert.fail();
        }

        // -----------------------
        // ■[異常系]チェック
        // -----------------------
        workbook = getWorkbook( path);

        Sheet sheet3 = workbook.getSheetAt( 2);

        parser = new ImageParamParser();

        reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( createTestData( ImageParamParser.DEFAULT_TAG));

        try {
            parseSheet( parser, sheet3, reportsParserInfo);
            fail( "チェックにかかっていない");
        } catch ( ParseException e) {
        }

        // -----------------------
        // □[正常系]必須パラメータがない場合
        // -----------------------
        workbook = getWorkbook();

        Sheet sheet4 = workbook.getSheetAt( 2);

        parser = new ImageParamParser();

        reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( createTestData( ImageParamParser.DEFAULT_TAG));

        try {
            parseSheet( parser, sheet4, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.getSheet( "Sheet4");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet4, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        }

        // ----------------------------------------------------------------------
        // □[正常系]1シート出力 / タグで複数画像 / 単一テンプレートを上書き
        // ----------------------------------------------------------------------
        workbook = getWorkbook();

        Sheet sheet5 = workbook.getSheetAt( INDEX_OF_SHEET5);

        parser = new ImageParamParser();

        reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( createPluralTestData( ImageParamParser.DEFAULT_TAG));

        try {
            parseSheet( parser, sheet5, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 不要シートの削除
        delSheetIndexs.clear();
        for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if ( i != INDEX_OF_SHEET5) {
                delSheetIndexs.add( i);
            }
        }
        for ( Integer i : delSheetIndexs) {
            workbook.removeSheetAt( i);
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.getSheet( "Sheet5");

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet5, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            // オブジェクトはチェックできないので、出力して確認
            String tmpDirPath = System.getProperty( "user.dir") + "/work/test/"; // System.getProperty( "java.io.tmpdir");
            File file = new File( tmpDirPath);
            if ( !file.exists()) {
                file.mkdirs();
            }

            try {
                String filepath = null;
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + "ImageParamParserTest3.xlsx";
                } else {
                    filepath = tmpDirPath + "ImageParamParserTest3.xls";
                }
                PoiUtil.writeBook( workbook, filepath);
                log.info( "出力ファイルを確認してください：" + filepath);

            } catch ( IOException e) {
                fail( e.toString());
            }
        }

        // ----------------------------------------------------------------------
        // □[正常系]1シート出力 / タグで複数画像 / 単一テンプレートを1回コピー
        // ----------------------------------------------------------------------

        workbook = getWorkbook();

        Sheet sheet6 = workbook.cloneSheet( INDEX_OF_SHEET5);

        parser = new ImageParamParser();

        reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setParamInfo( createPluralTestData( ImageParamParser.DEFAULT_TAG));

        try {
            parseSheet( parser, sheet6, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        // 不要シートの削除
        delSheetIndexs.clear();
        for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if ( i != workbook.getNumberOfSheets() - 1) {
                delSheetIndexs.add( i);
            }
        }
        for ( Integer i : delSheetIndexs) {
            workbook.removeSheetAt( i);
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        expectedSheet = expectedWorkbook.cloneSheet( INDEX_OF_SHEET5);

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, sheet6, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            // オブジェクトはチェックできないので、出力して確認
            String tmpDirPath = System.getProperty( "user.dir") + "/work/test/"; // System.getProperty( "java.io.tmpdir");
            File file = new File( tmpDirPath);
            if ( !file.exists()) {
                file.mkdirs();
            }

            try {
                String filepath = null;
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + "ImageParamParserTest4.xlsx";
                } else {
                    filepath = tmpDirPath + "ImageParamParserTest4.xls";
                }
                PoiUtil.writeBook( workbook, filepath);
                log.info( "出力ファイルを確認してください：" + filepath);

            } catch ( IOException e) {
                fail( e.toString());
            }
        }

        // ----------------------------------------------------------------------
        // □[正常系] 複数シート出力 / タグで複数画像 / 単一テンプレートを複数回コピー
        // ----------------------------------------------------------------------

        workbook = getWorkbook();

        for ( int i = 1; i <= PLURAL_COPY_FIRST_NUM_OF_SHEETS; i++) {

            Sheet sheet = workbook.cloneSheet( INDEX_OF_SHEET5);

            parser = new ImageParamParser();

            reportsParserInfo = new ReportsParserInfo();
            reportsParserInfo.setParamInfo( createPluralTestData( ImageParamParser.DEFAULT_TAG));

            try {
                parseSheet( parser, sheet, reportsParserInfo);
            } catch ( ParseException e) {
                fail( e.toString());
            }
        }

        // 不要シートの削除
        delSheetIndexs.clear();
        for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if ( i < workbook.getNumberOfSheets() - PLURAL_COPY_FIRST_NUM_OF_SHEETS) {
                delSheetIndexs.add( i);
            }
        }
        for ( Integer i : delSheetIndexs) {
            workbook.removeSheetAt( i);
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        for ( int i = 1; i <= PLURAL_COPY_FIRST_NUM_OF_SHEETS; i++) {
            expectedSheet = expectedWorkbook.cloneSheet( INDEX_OF_SHEET5);
        }

        try {
            int startOfTargetSheet = expectedWorkbook.getNumberOfSheets() - PLURAL_COPY_FIRST_NUM_OF_SHEETS;

            for ( int i = 0; i < PLURAL_COPY_FIRST_NUM_OF_SHEETS; i++) {
                // チェック
                ReportsTestUtil.checkSheet( expectedWorkbook.getSheetAt( startOfTargetSheet + i), workbook.getSheetAt( i), false);
            }
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            // オブジェクトはチェックできないので、出力して確認
            String tmpDirPath = System.getProperty( "user.dir") + "/work/test/"; // System.getProperty( "java.io.tmpdir");
            File file = new File( tmpDirPath);
            if ( !file.exists()) {
                file.mkdirs();
            }

            try {
                String filepath = null;
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + "ImageParamParserTest5.xlsx";
                } else {
                    filepath = tmpDirPath + "ImageParamParserTest5.xls";
                }
                PoiUtil.writeBook( workbook, filepath);
                log.info( "出力ファイルを確認してください：" + filepath);

            } catch ( IOException e) {
                fail( e.toString());
            }

        }

        // ----------------------------------------------------------------------
        // □[正常系] 複数シート出力 / タグで複数画像 / 複数(2個)のテンプレートを複数回コピー
        // ----------------------------------------------------------------------

        workbook = getWorkbook();

        Sheet sheet = null;
        int totalNumOfCopies = PLURAL_COPY_FIRST_NUM_OF_SHEETS + PLURAL_COPY_SECOND_NUM_OF_SHEETS;
        for ( int i = 1; i <= totalNumOfCopies; i++) {

            if ( i <= PLURAL_COPY_FIRST_NUM_OF_SHEETS) {
                sheet = workbook.cloneSheet( INDEX_OF_SHEET5);
            } else {
                sheet = workbook.cloneSheet( INDEX_OF_SHEET6);
            }

            parser = new ImageParamParser();

            reportsParserInfo = new ReportsParserInfo();
            reportsParserInfo.setParamInfo( createPluralTestData( ImageParamParser.DEFAULT_TAG));

            try {
                parseSheet( parser, sheet, reportsParserInfo);
            } catch ( ParseException e) {
                fail( e.toString());
            }
        }

        // 不要シートの削除
        delSheetIndexs.clear();
        for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if ( i < workbook.getNumberOfSheets() - totalNumOfCopies) {
                delSheetIndexs.add( i);
            }
        }
        for ( Integer i : delSheetIndexs) {
            workbook.removeSheetAt( i);
        }

        // 期待値ブックの読み込み
        expectedWorkbook = getExpectedWorkbook();
        for ( int i = 1; i <= totalNumOfCopies; i++) {
            if ( i <= PLURAL_COPY_FIRST_NUM_OF_SHEETS) {
                expectedSheet = expectedWorkbook.cloneSheet( INDEX_OF_SHEET5);
            } else {
                expectedSheet = expectedWorkbook.cloneSheet( INDEX_OF_SHEET6);
            }
        }

        try {
            int startOfTargetSheet = expectedWorkbook.getNumberOfSheets() - totalNumOfCopies;

            for ( int i = 0; i < totalNumOfCopies; i++) {
                // チェック
                ReportsTestUtil.checkSheet( expectedWorkbook.getSheetAt( startOfTargetSheet + i), workbook.getSheetAt( i), false);
            }
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            // オブジェクトはチェックできないので、出力して確認
            String tmpDirPath = System.getProperty( "user.dir") + "/work/test/"; // System.getProperty( "java.io.tmpdir");
            File file = new File( tmpDirPath);
            if ( !file.exists()) {
                file.mkdirs();
            }

            try {
                String filepath = null;
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + "ImageParamParserTest6.xlsx";
                } else {
                    filepath = tmpDirPath + "ImageParamParserTest6.xls";
                }
                PoiUtil.writeBook( workbook, filepath);
                log.info( "出力ファイルを確認してください：" + filepath);

            } catch ( IOException e) {
                fail( e.toString());
            }
        }

    }

    private ParamInfo createTestData( String tag) {

        ParamInfo info = new ParamInfo();
        info.addParam( tag, "png", getImagePath( "bbreak.PNG"));
        info.addParam( tag, "jpeg", getImagePath( "bbreak.JPG"));

        return info;

    }

    private ParamInfo createPluralTestData( String tag) {

        ParamInfo info = new ParamInfo();
        info.addParam( tag, "png", getImagePath( "bbreak.PNG"));
        info.addParam( tag, "jpeg", getImagePath( "bbreak.JPG"));

        info.addParam( tag, "expng", getImagePath( "excella.PNG"));
        info.addParam( tag, "exjpeg", getImagePath( "excella.JPG"));

        return info;

    }

    private String getImagePath( String fileName) {
        String path = null;

        try {
            URL url = this.getClass().getResource( fileName);
            path = URLDecoder.decode( url.getFile(), "UTF-8");

        } catch ( Exception e) {
            fail( e.toString());
        }

        return path;

    }

    /**
     * {@link org.bbreak.excella.reports.tag.ImageParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        // -----------------------
        // □制御行使用有無：無
        // -----------------------
        ImageParamParser parser = new ImageParamParser();
        assertFalse( parser.useControlRow());
    }

}
