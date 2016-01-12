package org.bbreak.excella.reports.samples;

import java.math.BigDecimal;
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;

/**
 * 交通費申請書サンプル出力クラス
 * 
 * @since 1.0
 */
public class TransExpenseReporter {

    private static Calendar calendar = Calendar.getInstance();
    
    /**
     * 交通費申請書サンプルの出力を実行します。
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
        String templateFileName = "交通費申請書テンプレート.xls";
        URL templateFileUrl = TransExpenseReporter.class.getResource( templateFileName);
        String templateFilePath = URLDecoder.decode( templateFileUrl.getPath(), "UTF-8");
        
        String outputFileName = "交通費申請書サンプル";
        String outputFileDir = "./";
        String outputFilePath = outputFileDir.concat( outputFileName);
        ReportBook outputBook = new ReportBook( templateFilePath, outputFilePath, ExcelExporter.FORMAT_TYPE);

        //
        // PDFにオプションを指定したい場合は
        // ConvertConfigurationを作成する
        //
//        ConvertConfiguration convertConfig = new ConvertConfiguration( OoPdfExporter.FORMAT_TYPE);
//        ReportBook outputBook = new ReportBook( templateFilePath, outputFilePath, convertConfig);
//        convertConfig.addOption( "RestrictPermissions", Boolean.TRUE); //制限設定：有
//        convertConfig.addOption( "PermissionPassword", "pass");        //編集パスワード
//        convertConfig.addOption( "Printing", 0);                       //印刷不可
//        convertConfig.addOption( "Changes", 4);                         //編集不可
        
        
        
        
        //
        // テンプレートファイル内のシート名と出力シート名を指定し、
        // ReportSheetインスタンスを生成して、ReportBookに追加します。
        //
        ReportSheet outputSheet = new ReportSheet( "申請書", "申請書");
        outputBook.addReportSheet( outputSheet);

        //
        // 置換パラメータをReportSheetオブジェクトに追加します。
        // (反復置換のパラメータには配列を渡します。)
        //
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "申請番号", "000001");

        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "申請日", calendar.getTime());

        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "社員番号", "000123");

        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "氏名", "ユーザ１");

        int targetYear = calendar.get( Calendar.YEAR);
        int targetMonth = calendar.get( Calendar.MONTH) - 1;
        calendar.set( targetYear, targetMonth, 1);
        outputSheet.addParam( SingleParamParser.DEFAULT_TAG, "対象年月", calendar.getTime());

        ParamInfo[] details = createDetailParamInfos( targetYear, targetMonth);
        outputSheet.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "明細", details);

        //
        // ReportProcessorインスタンスを生成し、
        // ReportBookを元にレポート処理を実行します。
        //
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.process( outputBook);

    }

    /**
     * 明細置換データを生成します。
     * 
     * @return 明細のListを返します。
     */
    private static ParamInfo[] createDetailParamInfos( int targetYear, int targetMonth) {

        List<ParamInfo> details = new ArrayList<ParamInfo>();
        ParamInfo detail = null;

        // 明細1行目
        detail = new ParamInfo();
        calendar.set( targetYear, targetMonth, 1);
        detail.addParam( SingleParamParser.DEFAULT_TAG, "日付", calendar.getTime());
        detail.addParam( SingleParamParser.DEFAULT_TAG, "乗車", "東京");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "降車", "品川");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "交通機関", "電車");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "摘要", "摘要A");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "金額", new BigDecimal( 160));
        details.add( detail);

        // 明細2行目
        detail = new ParamInfo();
        calendar.set( targetYear, targetMonth, 2);
        detail.addParam( SingleParamParser.DEFAULT_TAG, "日付", calendar.getTime());
        detail.addParam( SingleParamParser.DEFAULT_TAG, "乗車", "東京");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "降車", "新宿");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "交通機関", "電車");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "摘要", "摘要B");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "金額", new BigDecimal( 190));
        details.add( detail);

        // 明細3行目
        detail = new ParamInfo();
        calendar.set( targetYear, targetMonth, 3);
        detail.addParam( SingleParamParser.DEFAULT_TAG, "日付", calendar.getTime());
        detail.addParam( SingleParamParser.DEFAULT_TAG, "乗車", "東京");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "降車", "渋谷");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "", "電車");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "", "摘要C");
        detail.addParam( SingleParamParser.DEFAULT_TAG, "", new BigDecimal( 190));
        details.add( detail);

        return details.toArray( new ParamInfo[details.size()]);
    }

}
