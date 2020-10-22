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

package org.bbreak.excella.reports.others;

import java.io.File;
import java.io.FileInputStream;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLDecoder;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.processor.ReportsCheckException;
import org.bbreak.excella.reports.util.ReportsUtil;
import org.junit.jupiter.api.Test;

/**
 * テンプレートをシートコピーした時の印刷設定確認のためのテストクラス
 * 
 * @since 1.0
 */
public class PrintSettingCopyTest {

    /**
     * テンプレートシートのシートコピー数
     */
    private static final Integer SINGLE_COPY_NUM_OF_SHETTS = 3;

    /**
     * 複数テンプレートの場合の1枚目のテンプレートシートのコピー数
     */
    private static final Integer PLURAL_COPY_FIRST_NUM_OF_SHEETS = 3;

    /**
     * 複数テンプレートの場合の2枚目のテンプレートシートのコピー数
     */
    private static final Integer PLURAL_COPY_SECOND_NUM_OF_SHEETS = 4;

    /**
     * 印刷設定付のコピーをテストし、ファイルを出力します
     */
    @Test
    public void testCopyOfPrintSetting() throws Exception {
        // 印刷設定付テンプレートに上書きでファイル出力
        printWithPrintSetting( false);
        // 印刷設定付テンプレートをコピーしてファイル出力
        printWithPrintSetting( true);
        // 印刷設定付テンプレートを複数シートコピーしてファイル出力
        printPluralWithPrintSetting( true);
        // 印刷設定付テンプレート2シートを各々複数シートコピーしてファイル出力
        printPluralWithPrintSetting( false);
    }

    private static void printWithPrintSetting( boolean copyTemplate) throws Exception {
        // Excel 2003形式
        printExcel( "印刷設定有テンプレート.xls", copyTemplate);
        // Excel 2007形式
        printExcel( "印刷設定有テンプレート.xlsx", copyTemplate);
    }

    private static void printPluralWithPrintSetting( boolean singleTempSheet) throws Exception {
        // Excel 2003形式で出力
        if ( singleTempSheet) {
            printPluralExcel( "印刷設定有テンプレート.xls", singleTempSheet);
        } else {
            printPluralExcel( "印刷設定有複数テンプレート.xls", singleTempSheet);
        }
        // Excel 2007形式で出力
        if ( singleTempSheet) {
            printPluralExcel( "印刷設定有テンプレート.xlsx", singleTempSheet);
        } else {
            printPluralExcel( "印刷設定有複数テンプレート.xlsx", singleTempSheet);
        }
    }

    private static void printExcel( String templateFileName, boolean copyTemplate) throws Exception {
        // ReportBookの生成
        ReportBook outputBook = null;
        if ( copyTemplate) {
            outputBook = createReportBook( templateFileName, "Print_Setting_Copy_Template");
        } else {
            outputBook = createReportBook( templateFileName, "Print_Setting_Overwrite_Template");
        }
        // ReportSheetのセット
        setReportSheet( outputBook, copyTemplate);
        // シートコピー時の印刷設定の反映を確認
        Workbook workbook = getTemplateWorkbook( outputBook.getTemplateFileName());
        templateCopyTest( workbook, outputBook);
        // Processorの実行
        executeProcessor( new ReportProcessor(), outputBook);
    }

    private static void printPluralExcel( String templateFileName, boolean singleTempSheet) throws Exception {
        // ReportBookの生成
        ReportBook outputBook = null;
        if ( singleTempSheet) {
            outputBook = createReportBook( templateFileName, "Print_Setting_Single_Template_Plural_Copy");
        } else {
            outputBook = createReportBook( templateFileName, "Print_Setting_Plural_Template_Plural_Copy");
        }
        // ReportSheetのセット(複数)
        setReportSheets( singleTempSheet, outputBook);
        // シートコピー時の印刷設定の反映を確認
        Workbook workbook = getTemplateWorkbook( outputBook.getTemplateFileName());
        templateCopyTest( workbook, outputBook);
        // Processorの実行
        executeProcessor( new ReportProcessor(), outputBook);
    }

    private static void setReportSheet( ReportBook outputBook, boolean copyTemplate) throws Exception {
        ReportSheet outputSheet = null;
        if ( copyTemplate) {
            outputSheet = new ReportSheet( "請求書", "請求書コピー");
        } else {
            outputSheet = new ReportSheet( "請求書");
        }
        // ブックにシートをセット
        outputBook.addReportSheet( outputSheet);
    }

    private static void setReportSheets( boolean singleTempSheet, ReportBook outputBook) throws Exception {
        if ( singleTempSheet) {
            for ( int i = 1; i <= SINGLE_COPY_NUM_OF_SHETTS; i++) {
                ReportSheet outputSheet = new ReportSheet( "請求書", "請求書コピー" + i);
                // ブックにシートをセット
                outputBook.addReportSheet( outputSheet);
            }
        } else {
            for ( int i = 1; i <= PLURAL_COPY_FIRST_NUM_OF_SHEETS; i++) {
                ReportSheet outputSheet = new ReportSheet( "請求書A", "請求書Aコピー" + i);
                // ブックにシートをセット
                outputBook.addReportSheet( outputSheet);
            }
            int start = PLURAL_COPY_FIRST_NUM_OF_SHEETS + 1;
            int end = PLURAL_COPY_FIRST_NUM_OF_SHEETS + PLURAL_COPY_SECOND_NUM_OF_SHEETS;
            for ( int i = start; i <= end; i++) {
                ReportSheet outputSheet = new ReportSheet( "請求書B", "請求書Bコピー" + i);
                // ブックにシートをセット
                outputBook.addReportSheet( outputSheet);
            }
        }
    }

    private static void executeProcessor( ReportProcessor reportProcessor, ReportBook outputBook) throws Exception {
        reportProcessor.process( outputBook);
    }

    private static ReportBook createReportBook( String templateFileName, String outputFileName) throws UnsupportedEncodingException {
        String templateFilePath = getTemplateFilePath( templateFileName);
        String tmpDirPath = System.getProperty( "user.dir") + "/work/test/";
        File file = new File( tmpDirPath);
        if ( !file.exists()) {
            file.mkdirs();
        }
        String outputFilePath = tmpDirPath.concat( outputFileName);
        ReportBook outputBook = new ReportBook( templateFilePath, outputFilePath, ExcelExporter.FORMAT_TYPE);
        return outputBook;
    }

    private static String getTemplateFilePath( String templateFileName) throws UnsupportedEncodingException {
        URL templateFileUrl = PrintSettingCopyTest.class.getResource( templateFileName);
        String templateFilePath = URLDecoder.decode( templateFileUrl.getPath(), "UTF-8");
        return templateFilePath;
    }

    private static void templateCopyTest( Workbook workbook, ReportBook reportBook) throws ReportsCheckException {
        // 出力シート単位にコピーする
        for ( ReportSheet reportSheet : reportBook.getReportSheets()) {
            if ( reportSheet != null) {
                if ( !reportSheet.getSheetName().equals( reportSheet.getTemplateName())) {
                    // テンプレート名 !＝ 出力シート名
                    int tempIdx = workbook.getSheetIndex( reportSheet.getTemplateName());
                    Sheet expectedSheet = workbook.getSheetAt( tempIdx);
                    Sheet actualSheet = workbook.cloneSheet( tempIdx);
                    ReportsUtil.copyPrintSetup( workbook, tempIdx, actualSheet);
                    ReportsTestUtil.checkSheet( expectedSheet, actualSheet, true);
                }
            }
        }
    }

    private static Workbook getTemplateWorkbook( String filepath) throws Exception {
        FileInputStream fileIn = null;
        Workbook wb = null;

        try {
            fileIn = new FileInputStream( filepath);
            wb = WorkbookFactory.create( fileIn);
        } finally {
            if ( fileIn != null) {
                fileIn.close();
            }
        }

        return wb;
    }
}
