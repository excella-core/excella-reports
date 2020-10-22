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

package org.bbreak.excella.reports.exporter;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Date;

import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.junit.jupiter.api.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.ExcelOutputStreamExporter}のためのテスト・クラス。
 *
 * @since 1.1
 */
public class ExcelOutputStreamExporterTest {

    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();

    private OutputStream outputStream;

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelOutputStreamExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     *
     * @throws Exception
     */
    @Test
    public void testOutput() throws Exception {

        FileOutputStream xlsFileOutputStream = null;
        // 出力先
        String filePath = tmpDirPath + (new Date()).getTime();
        String xlsTemplateFileName = "ExcelOutputStreamExporterTest.xls";
        URL xlsTemplateFileUrl = ExcelOutputStreamExporterTest.class.getResource( xlsTemplateFileName);
        // XLSテスト� - フォーマット複数指定・サイズ比較テスト
        try {
            String xlsTemplateFilePath = URLDecoder.decode( xlsTemplateFileUrl.getPath(), "UTF-8");

            ReportBook xlsOutputBook = new ReportBook( xlsTemplateFilePath, filePath, ExcelExporter.FORMAT_TYPE, ExcelOutputStreamExporter.FORMAT_TYPE);

            ReportSheet outputSheet = new ReportSheet( "test");
            xlsOutputBook.addReportSheet( outputSheet);

            ReportProcessor xlsReportProcessor = new ReportProcessor();

            // Streamの処理
            File xlsStreamFile = new File( filePath + "_stream.xls");
            xlsFileOutputStream = new FileOutputStream( xlsStreamFile);

            xlsReportProcessor.addReportBookExporter( new ExcelOutputStreamExporter( xlsFileOutputStream));
            xlsReportProcessor.process( xlsOutputBook);

            File xlsExcelFile = new File( filePath + ".xls");

            // サイズ比較
            long xlsFileByteLength = xlsExcelFile.length();
            long xlsStreamFileSize = xlsStreamFile.length();
            assertEquals( xlsFileByteLength, xlsStreamFileSize);
        } finally {
            if ( xlsFileOutputStream != null) {
                xlsFileOutputStream.close();
            }
        }
        // XLSテスト� - フォーマット単数指定・ストリーム出力によるファイル以外が削除されていることを確認
        try {
            String xlsTemplateFilePath = URLDecoder.decode( xlsTemplateFileUrl.getPath(), "UTF-8");

            ReportBook xlsOutputBook = new ReportBook( xlsTemplateFilePath, filePath, ExcelOutputStreamExporter.FORMAT_TYPE);

            ReportSheet outputSheet = new ReportSheet( "test");
            xlsOutputBook.addReportSheet( outputSheet);

            ReportProcessor xlsReportProcessor = new ReportProcessor();

            // Streamの処理
            File xlsStreamFile = new File( filePath + "_exsistStream.xls");
            xlsFileOutputStream = new FileOutputStream( xlsStreamFile);

            xlsReportProcessor.addReportBookExporter( new ExcelOutputStreamExporter( xlsFileOutputStream));
            xlsReportProcessor.process( xlsOutputBook);

            File xlsExcelFile = new File( filePath + "_exsistsCheck.xls");

            // ストリーム出力されたファイルが開けること（0バイトでないファイルであることを確認）
            if ( xlsStreamFile.length() <= 0) {
                fail( "File is not opened");
            }

            // ファイル削除チェック
            if ( xlsExcelFile.exists()) {
                fail( "ExcelFile exists.");
            }
        } finally {
            if ( xlsFileOutputStream != null) {
                xlsFileOutputStream.close();
            }
        }

        // XLSX
        FileOutputStream xlsxFileOutputStream = null;
        try {

            String xlsxTemplateFileName = "ExcelOutputStreamExporterTest.xlsx";
            URL xlsxTemplateFileUrl = ExcelOutputStreamExporterTest.class.getResource( xlsxTemplateFileName);
            String xlsxTemplateFilePath = URLDecoder.decode( xlsxTemplateFileUrl.getPath(), "UTF-8");

            // TODO:XLSXの場合は複数指定不可
            ReportBook xlsxOutputBook = new ReportBook( xlsxTemplateFilePath, filePath, ExcelOutputStreamExporter.FORMAT_TYPE);

            ReportSheet outputSheet = new ReportSheet( "test");
            xlsxOutputBook.addReportSheet( outputSheet);

            // Stream書き込み用のファイル生成
            File xlsxStreamFile = new File( filePath + "_stream.xlsx");
            xlsxFileOutputStream = new FileOutputStream( xlsxStreamFile);

            // 処理実行
            ReportProcessor xlsxReportProcessor = new ReportProcessor();
            ExcelOutputStreamExporter exporter = new ExcelOutputStreamExporter( xlsxFileOutputStream);
            xlsxReportProcessor.addReportBookExporter( exporter);
            xlsxReportProcessor.process( xlsxOutputBook);

            // ストリーム出力されたファイルが開けること（0バイトでないファイルであることを確認）
            if ( xlsxStreamFile.length() <= 0) {
                fail( "File is not opened");
            }

        } finally {
            if ( xlsFileOutputStream != null) {
                xlsxFileOutputStream.close();
            }
        }
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelOutputStreamExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        ExcelOutputStreamExporter exporter = new ExcelOutputStreamExporter( outputStream);
        assertEquals( "OUTPUT_STREAM_EXCEL", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelOutputStreamExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        ExcelOutputStreamExporter exporter = new ExcelOutputStreamExporter( outputStream);
        assertEquals( "", exporter.getExtention());
    }
}
