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
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Map.Entry;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.tag.RowRepeatParamParser;

/**
 * ピボットグラフサンプル出力クラス
 * 
 * @since 1.0
 */
public class PivotGraphReporter {

    /**
     * ピボットグラフサンプルの出力を実行します。
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
        String templateFileName = "ピボットグラフテンプレート.xls";
        URL templateFileUrl = PivotGraphReporter.class.getResource( templateFileName);
        String templateFilePath = URLDecoder.decode( templateFileUrl.getPath(), "UTF-8");
        
        String outputFileName = "ピボットグラフサンプル";
        String outputFileDir = "./";
        String outputFilePath = outputFileDir.concat( outputFileName);
        ReportBook outputBook = new ReportBook( templateFilePath, outputFilePath, ExcelExporter.FORMAT_TYPE);

        //　
        // テンプレートファイル内のシート名と出力シート名を指定し、
        // ReportSheetインスタンスを生成して、ReportBookに追加します。
        outputBook.addReportSheet( new ReportSheet( "テーブル", "テーブル"));
        outputBook.addReportSheet( new ReportSheet( "グラフ", "グラフ"));
        ReportSheet outputDataSheet = new ReportSheet( "元データ", "元データ");
        outputBook.addReportSheet( outputDataSheet);

        //　
        // 置換パラメータをReportSheetオブジェクトに追加します。
        // (反復置換のパラメータには配列を渡します。)
        //　
        PivotGraphData[] datas = createRandomPivotGraphDatas();
        List<Date> salesDates = new ArrayList<Date>();
        List<String> productNames = new ArrayList<String>();
        List<BigDecimal> unitPrices = new ArrayList<BigDecimal>();
        List<BigDecimal> quantities = new ArrayList<BigDecimal>();
        List<BigDecimal> prices = new ArrayList<BigDecimal>();
        List<String> salesPersons = new ArrayList<String>();
        
        for(PivotGraphData data : datas) {
            salesDates.add( data.getSalesDate());
            productNames.add( data.getProductName());
            unitPrices.add( data.getUnitPrice());
            quantities.add( data.getQuantity());
            prices.add( data.getPrice());
            salesPersons.add( data.getSalesPerson());
        }
        
        outputDataSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "salesDate", salesDates.toArray());
        outputDataSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "productName", productNames.toArray());
        outputDataSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "unitPrice", unitPrices.toArray());
        outputDataSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "quantity", quantities.toArray());
        outputDataSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "price", prices.toArray());
        outputDataSheet.addParam( RowRepeatParamParser.DEFAULT_TAG, "salesPerson", salesPersons.toArray());
        
        // 
        // ReportProcessorインスタンスを生成し、
        // ReportBookを元にレポート処理を実行します。
        // 
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.process( outputBook);

    }

    /**
     * 売上データ置換パラメータを生成します(現在年月の月初から月末まで、数量・担当者ランダム指定)
     * 
     * @return 売上データ置換パラメータ(PivotGraphData配列)を返します
     */
    private static PivotGraphData[] createRandomPivotGraphDatas() throws Exception {

        List<PivotGraphData> datas = new ArrayList<PivotGraphData>();
        Random random = new Random();
        String[] salesPersons = new String[] {"販売員①", "販売員②", "販売員③"};

        Calendar calendar = Calendar.getInstance();
        int year = calendar.get( Calendar.YEAR);
        int month = calendar.get( Calendar.MONTH);
        int lastDayOfMonth = calendar.getActualMaximum( Calendar.DAY_OF_MONTH);

        Map<String, BigDecimal> productInfoMap = new LinkedHashMap<String, BigDecimal>();
        productInfoMap.put( "商品A", new BigDecimal( 10000));
        productInfoMap.put( "商品B", new BigDecimal( 9000));
        productInfoMap.put( "商品C", new BigDecimal( 7000));

        for ( int date = 1; date <= lastDayOfMonth; date++) {
            calendar.clear();
            calendar.set( year, month, date);

            for ( Entry<String, BigDecimal> entry : productInfoMap.entrySet()) {
                String productName = entry.getKey();
                BigDecimal unitPrice = entry.getValue();
                BigDecimal quantity = new BigDecimal( random.nextInt( 5) + 2);

                PivotGraphData data = new PivotGraphData();
                data.setSalesDate( calendar.getTime());
                data.setProductName( productName);
                data.setUnitPrice( unitPrice);
                data.setQuantity( quantity);
                data.setPrice( unitPrice.multiply( quantity));
                data.setSalesPerson( salesPersons[random.nextInt( 3)]);
                datas.add( data);
            }
        }

        return datas.toArray( new PivotGraphData[datas.size()]);
    }

}
