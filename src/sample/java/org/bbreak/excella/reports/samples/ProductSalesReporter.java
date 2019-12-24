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
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Map.Entry;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
import org.bbreak.excella.reports.tag.ColRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;

/**
 * 商品別売上実績サンプル出力クラス
 * 
 * @since 1.0
 */
public class ProductSalesReporter {

    private static Random random = new Random();

    /**
     * 商品別売上実績サンプル出力実行
     * 
     * @param args
     * @throws Exception
     */
    public static void main( String[] args) throws Exception {

        // 
        // ①読み込むテンプレートファイルのパス(拡張子含)
        // ②出力先のファイルパス(拡張子はExpoterによって自動的に付与されるため、不要。)
        // ③ファイルフォーマット(ConvertConfigurationの配列)
        // を指定し、ReportBookインスタンスを生成します。
        //
        String templateFileName = "商品売上実績テンプレート.xls";
        URL templateFileUrl = ProductSalesReporter.class.getResource( templateFileName);
        String templateFilePath = URLDecoder.decode( templateFileUrl.getPath(), "UTF-8");

        String outputFileDir = "./";
        String outputFileName1 = "商品売上実績(事業部Ⅰ)";
        String outputFileName2 = "商品売上実績(事業部Ⅱ)";
        String outputFilePath1 = outputFileDir.concat( outputFileName1);
        String outputFilePath2 = outputFileDir.concat( outputFileName2);
        ReportBook outputBook1 = new ReportBook( templateFilePath, outputFilePath1, ExcelExporter.FORMAT_TYPE);
        ReportBook outputBook2 = new ReportBook( templateFilePath, outputFilePath2, ExcelExporter.FORMAT_TYPE);

        Map<String, BigDecimal> productInfoMap = new LinkedHashMap<String, BigDecimal>();
        productInfoMap.put( "商品A", new BigDecimal( 10000));
        productInfoMap.put( "商品B", new BigDecimal( 9000));
        productInfoMap.put( "商品C", new BigDecimal( 7000));

        // 
        // ReportSheetインスタンスを生成して、ReportBookに追加します。
        // 
        List<ReportSheet> outputSheets1 = createOutputSheets( productInfoMap);
        outputBook1.addReportSheets( outputSheets1);
        List<ReportSheet> outputSheets2 = createOutputSheets( productInfoMap);
        outputBook2.addReportSheets( outputSheets2);

        // 
        // ReportProcessorインスタンスを生成し、
        // ReportBookを元にレポート処理を実行します。
        // 
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.process( outputBook1, outputBook2);

    }

    /**
     * ReportSheet(上期、下期の2シート)を生成します。
     * 
     * @return
     */
    private static List<ReportSheet> createOutputSheets( Map<String, BigDecimal> productInfoMap) {

        //
        // テンプレートファイル内のシート名と出力シート名を指定し、
        // ReportSheetインスタンスを生成します。
        // シートをコピーしたり、シート名を変更する場合はReportBook#setCopyTemplate()にTRUEを指定します。
        //
        String templateSheetName = "雛型";
        String outputSheetName1 = "上期";
        String outputSheetName2 = "下期";
        ReportSheet outputSheet1 = new ReportSheet( templateSheetName, outputSheetName1);
        ReportSheet outputSheet2 = new ReportSheet( templateSheetName, outputSheetName2);

        //
        // 置換パラメータをReportSheetオブジェクトに追加します。
        // (反復置換のパラメータには配列を渡します。)
        //
        String[] monthHeaders1 = {"4月", "5月", "6月", "第1期", "7月", "8月", "9月", "第2期"};
        outputSheet1.addParam( ColRepeatParamParser.DEFAULT_TAG, "対象月ヘッダ", monthHeaders1);
        String[] monthHeaders2 = {"10月", "11月", "12月", "第3期", "1月", "2月", "3月", "第4期"};
        outputSheet2.addParam( ColRepeatParamParser.DEFAULT_TAG, "対象月ヘッダ", monthHeaders2);

        // 商品データ
        List<ParamInfo> productDatas1 = new ArrayList<ParamInfo>();
        List<ParamInfo> productDatas2 = new ArrayList<ParamInfo>();
        for ( Entry<String, BigDecimal> entry : productInfoMap.entrySet()) {

            String productName = entry.getKey();
            BigDecimal unitPrice = entry.getValue();

            ParamInfo productData1 = createProductDataParamInfo( productName, unitPrice);
            productDatas1.add( productData1);
            ParamInfo productData2 = createProductDataParamInfo( productName, unitPrice);
            productDatas2.add( productData2);
        }
        outputSheet1.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "商品データ", productDatas1.toArray( new ParamInfo[productDatas1.size()]));
        outputSheet2.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "商品データ", productDatas2.toArray( new ParamInfo[productDatas2.size()]));

        List<ReportSheet> outputSheets = new ArrayList<ReportSheet>();
        outputSheets.add( outputSheet1);
        outputSheets.add( outputSheet2);

        return outputSheets;
    }

    /**
     * [商品データ]の置換パラメータを生成します。
     * 
     * @param productName
     * @param unitPrice
     * @return
     */
    private static ParamInfo createProductDataParamInfo( String productName, BigDecimal unitPrice) {

        ParamInfo productData = new ParamInfo();

        // 商品名
        productData.addParam( SingleParamParser.DEFAULT_TAG, "商品名", productName);

        // 四半期
        List<ParamInfo> quarterlyAmounts = new ArrayList<ParamInfo>();
        ParamInfo quarterlyAmount1st = createQuarterlyAmountParamInfo( unitPrice);
        quarterlyAmounts.add( quarterlyAmount1st);
        ParamInfo quarterlyAmount2nd = createQuarterlyAmountParamInfo( unitPrice);
        quarterlyAmounts.add( quarterlyAmount2nd);
        productData.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "四半期", quarterlyAmounts.toArray( new ParamInfo[quarterlyAmounts.size()]));

        return productData;
    }

    /**
     * [四半期]の置換パラメータを生成します。
     * 
     * @param unitPrice
     * @param monthFrom
     * @param monthTo
     * @return
     */
    private static ParamInfo createQuarterlyAmountParamInfo( BigDecimal unitPrice) {

        ParamInfo quarterlyAmount = new ParamInfo();

        // 月次
        List<ParamInfo> monthlyAmounts = new ArrayList<ParamInfo>();
        for ( int month = 1; month <= 3; month++) {
            ParamInfo monthlyAmount = createMonthlyAmountParamInfo( unitPrice);
            monthlyAmounts.add( monthlyAmount);
        }
        quarterlyAmount.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "月次", monthlyAmounts.toArray( new ParamInfo[monthlyAmounts.size()]));

        return quarterlyAmount;
    }

    /**
     * [月次]の置換パラメータを生成します。
     * 
     * @param unitPrice
     * @return
     */
    private static ParamInfo createMonthlyAmountParamInfo( BigDecimal unitPrice) {

        ParamInfo monthlyAmount = new ParamInfo();
        BigDecimal budget = unitPrice.multiply( new BigDecimal( random.nextInt( 5) + 5));
        BigDecimal actual = unitPrice.multiply( new BigDecimal( random.nextInt( 5) + 5));

        // 予算、実績、差額
        monthlyAmount.addParam( SingleParamParser.DEFAULT_TAG, "予算", budget);
        monthlyAmount.addParam( SingleParamParser.DEFAULT_TAG, "実績", actual);
        monthlyAmount.addParam( SingleParamParser.DEFAULT_TAG, "差額", actual.subtract( budget));

        return monthlyAmount;
    }

}
