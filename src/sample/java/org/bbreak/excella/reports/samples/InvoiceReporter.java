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

package org.bbreak.excella.reports.samples;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.tag.ImageParamParser;
import org.bbreak.excella.reports.tag.RowRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;

/**
 * 請求書サンプル出力クラス
 * 
 * @since 1.0
 */
public class InvoiceReporter {

    /**
     * 請求書サンプルの出力を実行します。
     * 
     * @param args
     * @throws Exception
     */
    public static void main( String[] args) throws Exception {

        //
        // ①読み込むテンプレートファイルのパス(拡張子含)
        // ②出力先のファイルパス(拡張子はExpoterによって自動的に付与されるため、不要。)
        // ③出力ファイルフォーマット(ConvertConfigurationの配列)
        // を指定し、ReportBookインスタンスを生成します。
        // 
        String templateFileName = "請求書テンプレート.xls";
        URL templateFileUrl = InvoiceReporter.class.getResource( templateFileName);
        String templateFilePath = URLDecoder.decode( templateFileUrl.getPath(), "UTF-8");

        String outputFileName = "請求書サンプル";
        String outputFileDir = "./";
        String outputFilePath = outputFileDir.concat( outputFileName);
        ReportBook outputBook = new ReportBook( templateFilePath, outputFilePath, ExcelExporter.FORMAT_TYPE);

        //　
        // テンプレートファイル内のシート名と出力シート名を指定し、
        // ReportSheetインスタンスを生成して、ReportBookに追加します。
        //　
        ReportSheet outputSheet = new ReportSheet( "請求書");
        outputBook.addReportSheet( outputSheet);

        //　
        // 置換パラメータをReportSheetオブジェクトに追加します。
        // (反復置換のパラメータには配列を渡します。)
        //　
        Calendar calendar = Calendar.getInstance();
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "請求日付", calendar.getTime());

        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "顧客名称", "○○商事　様");

        URL imageFileUrl = InvoiceReporter.class.getResource( "ロゴ.jpg");
        String imageFilePath = URLDecoder.decode( imageFileUrl.getPath(), "UTF-8");
        outputSheet.addParam( ImageParamParser.DEFAULT_TAG, "会社ロゴ", imageFilePath);

        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "差出人住所1", "〒100-0000");
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "差出人住所2", "東京都○○区○○○○ ×××ー×××");
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "差出人住所3", "×××ビル 9F");

        List<String> productNameList = new ArrayList<String>();
        productNameList.add( "商品A");
        productNameList.add( "商品B");
        productNameList.add( "商品C");
        outputSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "商品名", productNameList.toArray());

        List<BigDecimal> unitPriceList = new ArrayList<BigDecimal>();
        unitPriceList.add( new BigDecimal( 10000));
        unitPriceList.add( new BigDecimal( 9000));
        unitPriceList.add( new BigDecimal( 7000));
        outputSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "単価", unitPriceList.toArray());

        List<Integer> quantityList = new ArrayList<Integer>();
        quantityList.add( new Integer( 4));
        quantityList.add( new Integer( 5));
        quantityList.add( new Integer( 6));
        outputSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "数量", quantityList.toArray());

        List<BigDecimal> priceList = new ArrayList<BigDecimal>();
        priceList.add( new BigDecimal( 40000));
        priceList.add( new BigDecimal( 45000));
        priceList.add( new BigDecimal( 42000));
        outputSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "金額", priceList.toArray());

        BigDecimal subTotal = BigDecimal.ZERO;
        for ( BigDecimal price : priceList) {
            subTotal = subTotal.add( price);
        }
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "小計", subTotal);

        BigDecimal discount = new BigDecimal( 10000);
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "値引", discount);

        BigDecimal taxRate = new BigDecimal( 0.05);
        BigDecimal taxCharge = subTotal.multiply( taxRate).setScale( 0, RoundingMode.DOWN);
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "税額", taxCharge);

        BigDecimal total = subTotal.add( taxCharge).subtract( discount);
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "合計", total);

        // 
        // ReportProcessorインスタンスを生成し、
        // ReportBookを元にレポート処理を実行します。
        // 
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.process( outputBook);

    }

}
