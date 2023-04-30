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

package org.bbreak.excella.reports.processor;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.FileNotFoundException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.exporter.ReportBookExporter;
import org.bbreak.excella.reports.listener.PreBookParseListener;
import org.bbreak.excella.reports.listener.ReportProcessAdaptor;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.processor.ReportProcessor} のためのテスト・クラス。
 * 
 * @since 1.0
 */
@SuppressWarnings( "unchecked")
public class ReportProcessorTest extends ReportsWorkbookTest {

    /**
     * システム変数：テンプディレクトリ
     */
    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();

    /**
     * プロセスログ
     */
    private List<String> processStrings = new ArrayList<String>();

    /**
     * コンストラクタ
     * 
     * @param version バージョン
     */
    public ReportProcessorTest( String version) {
        super( version);
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#ReportProcessor()} のためのテスト・メソッド。
     */
    @Test
    public void testReportProcessor() {
        ReportProcessor processor = new ReportProcessor();
        // デフォルトタグパーサー
        Map<String, ReportsTagParser<?>> actualParsers = ( Map<String, ReportsTagParser<?>>) getPrivateField( processor, "parsers");
        Map<String, ReportsTagParser<?>> expectedParsers = ReportCreateHelper.createDefaultParsers();

        assertEquals( expectedParsers.size(), actualParsers.size());

        for ( Entry<String, ReportsTagParser<?>> entry : expectedParsers.entrySet()) {
            ReportsTagParser<?> actualParser = actualParsers.get( entry.getKey());
            assertNotNull( actualParser);
            assertEquals( entry.getValue().getClass(), actualParser.getClass());
        }

        // デフォルトエクスポーター
        Map<String, ReportBookExporter> actualExporters = ( Map<String, ReportBookExporter>) getPrivateField( processor, "exporters");
        Map<String, ReportBookExporter> expectedExporters = ReportCreateHelper.createDefaultExporters();

        assertEquals( expectedExporters.size(), actualExporters.size());

        for ( Entry<String, ReportBookExporter> entry : expectedExporters.entrySet()) {
            ReportBookExporter actualExporter = actualExporters.get( entry.getKey());
            assertNotNull( actualExporter);
            assertEquals( entry.getValue().getClass(), actualExporter.getClass());
        }

    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#process(org.bbreak.excella.reports.model.ReportData)} のためのテスト・メソッド。
     */
    @Test
    public void testProcess() {

        getWorkbook();

        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.addReportProcessListener( new CustomListener());

        reportProcessor.addReportsTagParser( new CustomTagParser( "$CUSTOM"));
        ReportBookExporter exporter = new CustomExporter();
        reportProcessor.addReportBookExporter( exporter);
        reportProcessor.addReportBookExporter( new ThrowExceptionExporter());

        // ------------------------------------
        // テンプレートコピーのない実行処理確認
        // ------------------------------------
        processStrings.clear();

        ConvertConfiguration configuration = new ConvertConfiguration( "CUSTOM");
        ConvertConfiguration[] configurations = new ConvertConfiguration[] {configuration, null};

        // ブック１
        ReportBook book1 = new ReportBook( getFilepath(), tmpDirPath + "workbookout1", configurations);
        // →シート１
        ReportSheet sheet1 = new ReportSheet( "Sheet1");
        book1.addReportSheet( sheet1);
        // →シート２
        ReportSheet sheet2 = new ReportSheet( "Sheet2");
        book1.addReportSheet( sheet2);

        // ブック２
        ReportBook book2 = new ReportBook( getFilepath(), tmpDirPath + "workbookout2", configurations);
        // →シート３
        ReportSheet sheet3 = new ReportSheet( "Sheet3");
        book2.addReportSheet( sheet3);

        book2.addReportSheet( null);

        try {
            reportProcessor.process( new ReportBook[] {book1, book2, null});
        } catch ( Exception e) {
            e.printStackTrace();
            fail( e.toString());
        }

        List<String> exceptedProccess = new ArrayList<String>();
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=Sheet1,シート2=Sheet2,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=Sheet3,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");

        assertArrayEquals( exceptedProccess.toArray(), processStrings.toArray());

        // 制御行の確認

        // ------------------------------------
        // テンプレートコピーのある実行処理確認
        // ------------------------------------
        processStrings.clear();

        // ブック１
        book1 = new ReportBook( getFilepath(), tmpDirPath + "workbookout1", configurations);
        // book1.setCopyTemplate( true);

        // →シート１
        ReportSheet sheetC1 = new ReportSheet( "Sheet1", "CSheet1");
        book1.addReportSheet( sheetC1);
        // →シート２
        ReportSheet sheetC2 = new ReportSheet( "Sheet2", "CSheet2");
        book1.addReportSheet( sheetC2);

        // ブック２
        book2 = new ReportBook( getFilepath(), tmpDirPath + "workbookout2", configurations);
        // →シート３
        ReportSheet sheetC3 = new ReportSheet( "Sheet3");
        book2.addReportSheet( sheetC3);

        book2.addReportSheet( null);

        try {
            reportProcessor.process( book1, book2, null);
        } catch ( Exception e) {
            e.printStackTrace();
            fail( e.toString());
        }

        exceptedProccess = new ArrayList<String>();
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=CSheet1,シート2=CSheet2,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=Sheet3,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");

        assertArrayEquals( exceptedProccess.toArray(), processStrings.toArray());

        // --------------------------------------------------------------------------
        // テンプレートファイルがない場合
        // --------------------------------------------------------------------------
        processStrings.clear();

        // ブック１
        book1 = new ReportBook( "test", tmpDirPath + "workbookout1", configurations);
        // book1.setCopyTemplate( true);

        // →シート１
        sheetC1 = new ReportSheet( "Sheet1", "CSheet1");
        book1.addReportSheet( sheetC1);

        try {
            reportProcessor.process( book1);
            fail();
        } catch ( FileNotFoundException e) {
            assertTrue( true);
        } catch ( Exception e) {
            e.printStackTrace();
            fail( e.toString());
        }
        assertTrue( processStrings.isEmpty());
        
        // ---------------------------------------------------------------------------
        // テンプレート上書き、 テンプレートコピー、利用しないテンプレートが混在する場合のシートオーダーの動作確認
        // ---------------------------------------------------------------------------
        processStrings.clear();

        // ブック１
        book1 = new ReportBook( getFilepath(), tmpDirPath + "workbookout1", configurations);

        // →シート１(コピー)
        ReportSheet sheetCopy1 = new ReportSheet( "Sheet1", "CSheet1");
        book1.addReportSheet( sheetCopy1);
        // →シート２（上書き）
        ReportSheet sheet2Original = new ReportSheet( "Sheet2");
        book1.addReportSheet( sheet2Original);
        // →シート２（コピー）
        ReportSheet sheetCopy2 = new ReportSheet( "Sheet2", "CSheet2");
        book1.addReportSheet( sheetCopy2);
        // →シート３（上書き）
        ReportSheet sheet3Original = new ReportSheet( "Sheet3");
        book1.addReportSheet( sheet3Original);
        // →シート４（出力対象外）

        // ブック２
        book2 = new ReportBook( getFilepath(), tmpDirPath + "workbookout2", configurations);
        // →シート１(出力対象外)
        // →シート２(コピー）
        ReportSheet sheet2Copy = new ReportSheet( "Sheet2", "CSheet2");
        book2.addReportSheet( sheet2Copy);
        // →シート２（上書き）
        ReportSheet sheetOriginal2 = new ReportSheet( "Sheet2");
        book2.addReportSheet( sheetOriginal2);
        // →シート３(コピー）
        ReportSheet sheet3Copy = new ReportSheet( "Sheet3", "CSheet3");
        book2.addReportSheet( sheet3Copy);
        // →シート４（上書き）
        ReportSheet sheetOriginal3 = new ReportSheet( "Sheet3");
        book2.addReportSheet( sheetOriginal3);
        // →シート４（出力対象外）

        try {
            reportProcessor.process( book1, book2);
        } catch ( Exception e) {
            e.printStackTrace();
            fail( e.toString());
        }

        exceptedProccess = new ArrayList<String>();
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=CSheet1,シート2=Sheet2,シート3=CSheet2,シート4=Sheet3,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=CSheet2,シート2=Sheet2,シート3=CSheet3,シート4=Sheet3,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");

        assertArrayEquals( exceptedProccess.toArray(), processStrings.toArray());

        // --------------------------------------------------------------------------
        // Exportの出力処理時にエラー発生の場合の出力後処理の動作確認
        // --------------------------------------------------------------------------
        // 出力時のエラー後、出力後処理
        processStrings.clear();

        ConvertConfiguration configuration2 = new ConvertConfiguration( "E");
        configurations = new ConvertConfiguration[] {configuration, configuration2};

        // ブック１
        book1 = new ReportBook( getFilepath(), tmpDirPath + "workbookout1", configurations);
        // book1.setCopyTemplate( true);

        // →シート１
        sheetC1 = new ReportSheet( "Sheet1", "CSheet1");
        book1.addReportSheet( sheetC1);

        try {
            reportProcessor.process( book1);
            fail();
        } catch ( Exception e) {
            assertTrue( true);
        }

        exceptedProccess = new ArrayList<String>();
        exceptedProccess.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
        exceptedProccess.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
        exceptedProccess.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
        exceptedProccess.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");
        exceptedProccess.add( "ブック処理結果:シート1=CSheet1,");
        exceptedProccess.add( "出力前処理 CustomExporter#setup()");
        exceptedProccess.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 CustomExporter#tearDown()");
        exceptedProccess.add( "出力前処理 ThrowExceptionExporter#setup()");
        exceptedProccess.add( "出力処理 ThrowExceptionExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        exceptedProccess.add( "出力後処理 ThrowExceptionExporter#tearDown()");

        assertArrayEquals( exceptedProccess.toArray(), processStrings.toArray());

    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#addReportsTagParser(org.bbreak.excella.reports.tag.ReportsTagParser)} のためのテスト・メソッド。
     */
    @Test
    public void testAddReportsTagParser() {
        ReportProcessor processor = new ReportProcessor();

        ReportsTagParser<?> addparser = new CustomTagParser( "$CUSTOM");
        processor.addReportsTagParser( addparser);

        Map<String, ReportsTagParser<?>> actualParsers = ( Map<String, ReportsTagParser<?>>) getPrivateField( processor, "parsers");
        assertEquals( addparser, actualParsers.get( "$CUSTOM"));
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#removeReportsTagParser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveReportsTagParser() {

        ReportProcessor processor = new ReportProcessor();

        processor.removeReportsTagParser( "");
        Map<String, ReportsTagParser<?>> actualParsers = ( Map<String, ReportsTagParser<?>>) getPrivateField( processor, "parsers");

        assertNull( actualParsers.get( ""));
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#clearReportsTagParser()} のためのテスト・メソッド。
     */
    @Test
    public void testClearReportsTagParser() {
        ReportProcessor processor = new ReportProcessor();
        processor.clearReportsTagParser();
        Map<String, ReportsTagParser<?>> actualParsers = ( Map<String, ReportsTagParser<?>>) getPrivateField( processor, "parsers");

        assertTrue( actualParsers.isEmpty());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#addReportBookExporter(org.bbreak.excella.reports.exporter.ReportBookExporter)} のためのテスト・メソッド。
     */
    @Test
    public void testAddReportBookExporter() {
        ReportProcessor processor = new ReportProcessor();

        ReportBookExporter bookExporter = new CustomExporter();
        processor.addReportBookExporter( bookExporter);

        Map<String, ReportBookExporter> actualExporters = ( Map<String, ReportBookExporter>) getPrivateField( processor, "exporters");

        assertEquals( bookExporter, actualExporters.get( "CUSTOM"));
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#removeReportBookExporter(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveReportBookExporter() {
        ReportProcessor processor = new ReportProcessor();

        processor.removeReportBookExporter( "EXCEL");

        Map<String, ReportBookExporter> actualExporters = ( Map<String, ReportBookExporter>) getPrivateField( processor, "exporters");

        assertNull( actualExporters.get( "EXCEL"));
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#clearReportBookExporter()} のためのテスト・メソッド。
     */
    @Test
    public void testClearReportBookExporter() {
        ReportProcessor processor = new ReportProcessor();

        processor.clearReportBookExporter();

        Map<String, ReportBookExporter> actualExporters = ( Map<String, ReportBookExporter>) getPrivateField( processor, "exporters");

        assertTrue( actualExporters.isEmpty());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#addReportProcessListener(org.bbreak.excella.reports.listener.ReportProcessListener)} のためのテスト・メソッド。
     */
    @Test
    public void testAddReportProcessListener() {
        ReportProcessor processor = new ReportProcessor();
        CustomListener listener = new CustomListener();

        ReportProcessListenerContainer listeners = processor.listeners;
        List<PreBookParseListener> listenerList = ( List<PreBookParseListener>) getPrivateField( listeners, "preParseListeners");

        assertTrue( listenerList.isEmpty());

        processor.addReportProcessListener( listener);

        assertEquals( 1, listenerList.size());
        assertEquals( listener, listenerList.iterator().next());

    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportProcessor#removeReportProcessListener(org.bbreak.excella.reports.listener.ReportProcessListener)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveReportProcessListener() {

        ReportProcessor processor = new ReportProcessor();
        CustomListener listener = new CustomListener();
        processor.addReportProcessListener( listener);

        ReportProcessListenerContainer listeners = processor.listeners;
        List<PreBookParseListener> listenerList = ( List<PreBookParseListener>) getPrivateField( listeners, "preParseListeners");

        assertEquals( 1, listenerList.size());

        processor.removeReportProcessListener( listener);

        assertTrue( listenerList.isEmpty());

    }

    private Object getPrivateField( Object object, String fieldName) {

        Field field = null;
        Object res = null;
        try {
            field = object.getClass().getDeclaredField( fieldName);
            field.setAccessible( true);
            res = field.get( object);
        } catch ( Exception e) {
            e.printStackTrace();
            fail();
        }
        return res;

    }

    class CustomTagParser extends ReportsTagParser<Object> {

        public CustomTagParser( String tag) {
            super( tag);
        }

        @Override
        public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            processStrings.add( "タグ解析処理 CustomTagParser#parse( Sheet sheet, Cell tagCell, Object data)");
            return null;
        }

        @Override
        public boolean useControlRow() {
            return false;
        }

    }

    class CustomExporter extends ReportBookExporter {

        @Override
        public String getExtention() {
            return ".custom";
        }

        @Override
        public String getFormatType() {
            return "CUSTOM";
        }

        @Override
        public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
            processStrings.add( "出力処理 CustomExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
        }

        @Override
        public void setup() {
            processStrings.add( "出力前処理 CustomExporter#setup()");
            super.setup();
        }

        @Override
        public void tearDown() {
            processStrings.add( "出力後処理 CustomExporter#tearDown()");
            super.tearDown();
        }

    }

    class ThrowExceptionExporter extends ReportBookExporter {

        @Override
        public String getExtention() {
            return ".e";
        }

        @Override
        public String getFormatType() {
            return "E";
        }

        @Override
        public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
            processStrings.add( "出力処理 ThrowExceptionExporter#output( Workbook book, BookData bookdata, ConvertConfiguration configuration)");
            throw new ExportException( new Exception());
        }

        @Override
        public void setup() {
            processStrings.add( "出力前処理 ThrowExceptionExporter#setup()");
            super.setup();
        }

        @Override
        public void tearDown() {
            processStrings.add( "出力後処理 ThrowExceptionExporter#tearDown()");
            super.tearDown();
        }

    }

    class CustomListener extends ReportProcessAdaptor {

        @Override
        public void postBookParse( Workbook workbook, ReportBook reportBook) {
            processStrings.add( "ブック解析後処理 CustomListener#postBookParse( Workbook workbook, ReportBook reportBook)");

            StringBuffer buffer = new StringBuffer( "ブック処理結果:");
            for ( int i = 0; i < workbook.getNumberOfSheets(); i++) {
                buffer.append( "シート" + (i + 1) + "=" + workbook.getSheetName( i) + ",");
            }
            processStrings.add( buffer.toString());
            System.out.println( buffer.toString());

            // System.out.println("workbook=" + workbook + ",reportBook=" + reportBook);
            super.postBookParse( workbook, reportBook);
        }

        @Override
        public void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException {
            processStrings.add( "シート解析後処理 CustomListener#postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData)");
            // System.out.println("sheet=" + sheet + ",sheetParser=" + sheetParser+ ",sheetData=" + sheetData);
            super.postParse( sheet, sheetParser, sheetData);
        }

        @Override
        public void preBookParse( Workbook workbook, ReportBook reportBook) {
            processStrings.add( "ブック解析前処理 CustomListener#preBookParse( Workbook workbook, ReportBook reportBook)");
            // System.out.println("workbook=" + workbook +",reportBook=" + reportBook);
            super.preBookParse( workbook, reportBook);
        }

        @Override
        public void preParse( Sheet sheet, SheetParser sheetParser) throws ParseException {
            processStrings.add( "シート解析前処理 CustomListener#preParse( Sheet sheet, SheetParser sheetParser)");
            // System.out.println("sheet=" + sheet +",SheetParser=" + sheetParser);
            super.preParse( sheet, sheetParser);
        }

    }

}
