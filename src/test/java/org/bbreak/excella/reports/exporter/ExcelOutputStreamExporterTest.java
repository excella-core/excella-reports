/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ExcelOutputStreamExporterTest.java 80 2009-10-30 02:29:59Z a-hoshino $
 * $Revision: 80 $
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
package org.bbreak.excella.reports.exporter;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Date;

import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.junit.Test;

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
    public void testOutput() {

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
        } catch ( Exception e) {
            e.getStackTrace();
            fail( e.toString());
        } finally {
            try {
                if ( xlsFileOutputStream != null) {
                    xlsFileOutputStream.close();
                }
            } catch ( IOException e) {
                e.printStackTrace();
                fail( e.toString());
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
        } catch ( Exception e) {
            e.getStackTrace();
            fail( e.toString());
        } finally {
            try {
                if ( xlsFileOutputStream != null) {
                    xlsFileOutputStream.close();
                }
            } catch ( IOException e) {
                e.printStackTrace();
                fail( e.toString());
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

        } catch ( Exception e) {
            e.getStackTrace();
            fail( e.toString());
        } finally {
            try {
                if ( xlsFileOutputStream != null) {
                    xlsxFileOutputStream.close();
                }
            } catch ( IOException e) {
                e.printStackTrace();
                fail( e.toString());
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
