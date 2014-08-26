package org.bbreak.excella.reports.others;

import java.io.File;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLDecoder;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.tag.ImageParamParser;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * 画像描画確認のためのテストクラス
 * 
 * @since 1.0
 */
public class ImageDisplayTest {

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
     * 目視確認用の画像付ファイルを出力します
     */
    @Test
    public void testImagePrint() throws Exception {
        // 画像タグ付テンプレートに上書きで帳票出力
        printTagedImage( false);
        // 画像タグ付テンプレートをコピーして帳票出力
        printTagedImage( true);

        // 画像付テンプレートに上書きで帳票出力
        printImage( false);
        // 画像付テンプレートをコピーして帳票出力
        printImage( true);

        // 画像タグ付テンプレートを複数シートコピーして帳票出力
        printPluralTagedImage( true);
        // 画像タグ付テンプレート2シートを各々複数シートコピーして帳票出力
        printPluralTagedImage( false);

        // 画像付テンプレートを複数シートコピーして帳票出力
        printPluralImage( true);
        // 画像付テンプレート2シートを各々複数シートコピーして帳票出力
        printPluralImage( false);
    }

    private static void printTagedImage( boolean copyTemplate) throws Exception {
        boolean useTag = true;
        // Excel 2003形式で出力
        printExcel( "画像タグ付テンプレート.xls", copyTemplate, useTag);
        // Excel 2007形式で出力
        printExcel( "画像タグ付テンプレート.xlsx", copyTemplate, useTag);
    }

    private static void printImage( boolean copyTemplate) throws Exception {
        boolean useTag = false;
        // Excel 2003形式で出力
        printExcel( "画像付テンプレート.xls", copyTemplate, useTag);
        // Excel 2007形式で出力
        printExcel( "画像付テンプレート.xlsx", copyTemplate, useTag);

    }

    private static void printPluralTagedImage( boolean singleTempSheet) throws Exception {
        boolean useTag = true;
        // Excel 2003形式で出力
        if ( singleTempSheet) {
            printPluralExcel( "画像タグ付テンプレート.xls", singleTempSheet, useTag);
        } else {
            printPluralExcel( "画像タグ付複数テンプレート.xls", singleTempSheet, useTag);
        }
        // Excel 2007形式で出力
        if ( singleTempSheet) {
            printPluralExcel( "画像タグ付テンプレート.xlsx", singleTempSheet, useTag);
        } else {
            printPluralExcel( "画像タグ付複数テンプレート.xlsx", singleTempSheet, useTag);
        }
    }

    private static void printPluralImage( boolean singleTempSheet) throws Exception {
        boolean useTag = false;
        // Excel 2003形式で出力
        if ( singleTempSheet) {
            printPluralExcel( "画像付テンプレート.xls", singleTempSheet, useTag);
        } else {
            printPluralExcel( "画像付複数テンプレート.xls", singleTempSheet, useTag);
        }
        // Excel 2007形式で出力
        if ( singleTempSheet) {
            printPluralExcel( "画像付テンプレート.xlsx", singleTempSheet, useTag);
        } else {
            printPluralExcel( "画像付複数テンプレート.xlsx", singleTempSheet, useTag);
        }
    }

    private static void printExcel( String templateFileName, boolean copyTemplate, boolean useTag) throws Exception {
        // ReportBookの生成
        ReportBook outputBook = null;
        if ( copyTemplate) {
            if ( useTag) {
                outputBook = createReportBook( templateFileName, "TagImage_Single_Template_Single_Copy");
            } else {
                outputBook = createReportBook( templateFileName, "Image_Single_Template_Single_Copy");
            }
        } else {
            if ( useTag) {
                outputBook = createReportBook( templateFileName, "TagImage_OverWrite_Template");
            } else {
                outputBook = createReportBook( templateFileName, "Image_OverWrite_Template");
            }
        }
        // ReportSheetのセット
        setReportSheet( outputBook, copyTemplate);
        // Processorの実行
        executeProcessor( new ReportProcessor(), outputBook);
    }

    private static void printPluralExcel( String templateFileName, boolean singleTempSheet, boolean useTag) throws Exception {
        // ReportBookの生成
        ReportBook outputBook = null;
        if ( singleTempSheet) {
            if ( useTag) {
                outputBook = createReportBook( templateFileName, "TagImage_Single_Template_Plural_Copy");
            } else {
                outputBook = createReportBook( templateFileName, "Image_Single_Template_Plural_Copy");
            }
        } else {
            if ( useTag) {
                outputBook = createReportBook( templateFileName, "TagImage_Plural_Template_Plural_Copy");
            } else {
                outputBook = createReportBook( templateFileName, "Image_Plural_Template_Plural_Copy");
            }
        }
        // ReportSheetのセット(複数)
        setReportSheets( singleTempSheet, outputBook);
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
        // シートにパラメータを設定
        setParamToSheet( outputSheet);
        // ブックにシートをセット
        outputBook.addReportSheet( outputSheet);
    }

    private static void setReportSheets( boolean singleTempSheet, ReportBook outputBook) throws Exception {
        if ( singleTempSheet) {
            for ( int i = 1; i <= SINGLE_COPY_NUM_OF_SHETTS; i++) {
                ReportSheet outputSheet = new ReportSheet( "請求書", "請求書コピー" + i);
                // シートにパラメータを設定
                setParamToSheet( outputSheet);
                // ブックにシートをセット
                outputBook.addReportSheet( outputSheet);
            }
        } else {
            for ( int i = 1; i <= PLURAL_COPY_FIRST_NUM_OF_SHEETS; i++) {
                ReportSheet outputSheet = new ReportSheet( "請求書A", "請求書Aコピー" + i);
                // シートにパラメータを設定
                setParamToSheet( outputSheet);
                // ブックにシートをセット
                outputBook.addReportSheet( outputSheet);
            }
            int start = PLURAL_COPY_FIRST_NUM_OF_SHEETS + 1;
            int end = PLURAL_COPY_FIRST_NUM_OF_SHEETS + PLURAL_COPY_SECOND_NUM_OF_SHEETS;
            for ( int i = start; i <= end; i++) {
                ReportSheet outputSheet = new ReportSheet( "請求書B", "請求書Bコピー" + i);
                // シートにパラメータを設定
                setParamToSheet( outputSheet);
                // ブックにシートをセット
                outputBook.addReportSheet( outputSheet);
            }
        }
    }

    private static void executeProcessor( ReportProcessor reportProcessor, ReportBook outputBook) throws Exception {
        reportProcessor.process( outputBook);
    }

    private static void setParamToSheet( ReportSheet outputSheet) throws Exception {
        URL imageFileUrl = ImageDisplayTest.class.getResource( "ロゴ.JPG");
        String imageFilePath = URLDecoder.decode( imageFileUrl.getPath(), "UTF-8");
        outputSheet.addParam( ImageParamParser.DEFAULT_TAG, "会社ロゴ", imageFilePath);

        URL imageFileUrl2 = ImageDisplayTest.class.getResource( "ロゴ２.PNG");
        String imageFilePath2 = URLDecoder.decode( imageFileUrl2.getPath(), "UTF-8");
        outputSheet.addParam( ImageParamParser.DEFAULT_TAG, "会社ロゴ２", imageFilePath2);
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
        URL templateFileUrl = ImageDisplayTest.class.getResource( templateFileName);
        String templateFilePath = URLDecoder.decode( templateFileUrl.getPath(), "UTF-8");
        return templateFilePath;
    }

}
