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

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Optional;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.bbreak.excella.core.CellClone;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.util.ReportsUtil;

/**
 * シート内のイメージ置換文字列を変換するパーサ
 * 
 * @since 1.0
 */
public class ImageParamParser extends ReportsTagParser<String> {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( ImageParamParser.class);

    /**
     * JPEGファイル末尾
     */
    private static final String JPEG_SUFFIX = ".jpeg";

    /**
     * JPGファイル末尾
     */
    private static final String JPG_SUFFIX = ".jpg";

    /**
     * PNGファイル末尾
     */
    private static final String PNG_SUFFIX = ".png";

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$I";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * 画像の幅調整用のパラメータ
     */
    public static final String PARAM_WIDTH = "widthAdjustment";

    /**
     * 画像の高さ調整用のパラメータ
     */
    public static final String PARAM_HEIGHT = "heightAdjustment";

    /**
     * 画像サイズ倍率用のパラメータ
     */
    public static final String PARAM_SCALE = "scale";

    /**
     * リサイズ基準指定用のパラメータ(セル基準: {@code "cell"}, 元画像基準: {@code "image"})
     */
    public static final String PARAM_RESIZE_BASE = "resizeBase";

    private enum ResizeBase {
        CELL( "cell"), IMAGE( "image");

        private static final ResizeBase DEFAULT = IMAGE;

        private final String code;

        ResizeBase( String code) {
            this.code = code;
        }

        static ResizeBase ofCode( String code) {
            Optional<ResizeBase> base = Arrays.stream( values())
                .filter( b -> b.code.equalsIgnoreCase( code))
                .findAny();
            return base.orElse( DEFAULT);
        }
    }

    /**
     * 図オブジェクトコンテナのキャッシュ
     */
    @SuppressWarnings( "rawtypes")
    private Map<Sheet, Drawing> drawingCash = new HashMap<Sheet, Drawing>();

    /**
     * コンストラクタ
     */
    public ImageParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ImageParamParser( String tag) {
        super( tag);
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

        // コメント有無チェック
        if ( hasComments( sheet)) {
            throw new ParseException( "不正シート[" + sheet.getWorkbook().getSheetName( sheet.getWorkbook().getSheetIndex( sheet)) + "]：コメント有");
        }

        // パラメータチェック
        checkParam( paramDef, tagCell);

        ReportsParserInfo reportsParserInfo = ( ReportsParserInfo) data;

        ParamInfo paramInfo = reportsParserInfo.getParamInfo();

        String paramValue = null;

        if ( paramInfo != null) {
            // 置換
            String replaceParam = paramDef.get( PARAM_VALUE);

            paramValue = getParamData( paramInfo, replaceParam);

            // 画像の幅調整値
            Integer dx1 = null;
            if ( paramDef.containsKey( PARAM_WIDTH)) {
                dx1 = Integer.valueOf( paramDef.get( PARAM_WIDTH));
            }

            // 画像の高さ調整値
            Integer dy1 = null;
            if ( paramDef.containsKey( PARAM_HEIGHT)) {
                dy1 = Integer.valueOf( paramDef.get( PARAM_HEIGHT));
            }

            // 画像の倍率
            Double scale = 1.0;
            if ( paramDef.containsKey( PARAM_SCALE)) {
                scale = Double.valueOf( paramDef.get( PARAM_SCALE));
            }

            boolean inMergedRegion = ReportsUtil.getMergedAddress( sheet, tagCell.getRowIndex(), tagCell.getColumnIndex()) != null;

            // リサイズ基準
            ResizeBase resizeBase = inMergedRegion ? ResizeBase.CELL : ResizeBase.DEFAULT;
            if ( paramDef.containsKey( PARAM_RESIZE_BASE)) {
                resizeBase = ResizeBase.ofCode(paramDef.get( PARAM_RESIZE_BASE));
            }

            // 結合セルに含まれるか
            if ( inMergedRegion) {
                CellStyle cellStyle = tagCell.getCellStyle();
                tagCell.setBlank();
                tagCell.setCellStyle( cellStyle);
            } else {
                tagCell = new CellClone( tagCell);
                PoiUtil.clearCell( sheet, new CellRangeAddress( tagCell.getRowIndex(), tagCell.getRowIndex(), tagCell.getColumnIndex(), tagCell.getColumnIndex()));
            }

            if ( log.isDebugEnabled()) {
                Workbook workbook = sheet.getWorkbook();
                String sheetName = workbook.getSheetName( workbook.getSheetIndex( sheet));

                log.debug( "[シート名=" + sheetName + ",セル=(" + tagCell.getRowIndex() + "," + tagCell.getColumnIndex() + ")]  " + tagCell.getStringCellValue() + " ⇒ " + paramValue);
            }

            if ( paramValue != null) {
                replaceImageValue( sheet, tagCell, paramValue, dx1, dy1, scale, resizeBase);
            }

        }

        // 解析結果の生成
        ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
        parsedReportInfo.setParsedObject( paramValue);
        parsedReportInfo.setRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setColumnIndex( tagCell.getColumnIndex());
        parsedReportInfo.setDefaultRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setDefaultColumnIndex( tagCell.getColumnIndex());

        return parsedReportInfo;
    }

    /**
     * 不正なパラメータがある場合、ParseExceptionをthrowする。
     * 
     * @param paramDef パラメータ定義
     * @param tagCell タグセル
     * @throws ParseException 不正なパラメータがある場合
     */
    private void checkParam( Map<String, String> paramDef, Cell tagCell) throws ParseException {
        // 画像の幅と高さ調整値が整数以外は不可
        if ( paramDef.containsKey( PARAM_WIDTH)) {
            try {
                Integer.parseInt( paramDef.get( PARAM_WIDTH));
            } catch ( NumberFormatException e) {
                throw new ParseException( tagCell, "整数以外の値" + PARAM_WIDTH + ":" + paramDef.get( PARAM_WIDTH), e);
            }
        }
        if ( paramDef.containsKey( PARAM_HEIGHT)) {
            try {
                Integer.parseInt( paramDef.get( PARAM_HEIGHT));
            } catch ( NumberFormatException e) {
                throw new ParseException( tagCell, "整数以外の値" + PARAM_HEIGHT + ":" + paramDef.get( PARAM_HEIGHT), e);
            }
        }
        // 画像の倍率値は実数以外は不可
        if ( paramDef.containsKey( PARAM_SCALE)) {
            try {
                Double.parseDouble( paramDef.get( PARAM_SCALE));
            } catch ( NumberFormatException e) {
                throw new ParseException( tagCell, "実数以外の値" + PARAM_SCALE + ":" + paramDef.get( PARAM_SCALE), e);
            }
        }

    }

    /**
     * シートにコメントが設定されているかを取得する。
     * 
     * @param sheet シート
     * @return 有：true／無：false
     */
    private boolean hasComments( Sheet sheet) {
        if ( sheet instanceof XSSFSheet) {
            XSSFSheet xssfSheet = ( XSSFSheet) sheet;
            return xssfSheet.hasComments();
        } else if ( sheet instanceof HSSFSheet) {
            Iterator<Row> rowIterator = sheet.iterator();
            while ( rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.iterator();
                while ( cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    Comment comment = cell.getCellComment();
                    if ( comment != null) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * セルのパラメータ置換
     * 
     * @param sheet 処理シート
     * @param cell タグセル
     * @param filePath 置換イメージファイルパス
     * @param dx1 画像の幅調整値
     * @param dy1 画像の高さ調整値
     * @param scale 画像の倍率値
     * @param resizeBase リサイズ基準
     * @throws ParseException
     */
    public void replaceImageValue( Sheet sheet, Cell cell, String filePath, Integer dx1, Integer dy1, Double scale, ResizeBase resizeBase) throws ParseException {

        Workbook workbook = sheet.getWorkbook();

        int format = -1;
        if ( filePath.toLowerCase().endsWith( JPEG_SUFFIX) || filePath.toLowerCase().endsWith( JPG_SUFFIX)) {
            format = Workbook.PICTURE_TYPE_JPEG;
        } else if ( filePath.toLowerCase().endsWith( PNG_SUFFIX)) {
            format = Workbook.PICTURE_TYPE_PNG;
        }
        if ( format == -1) {
            throw new ParseException( cell, "未対応の画像フォーマットファイルが指定されています。" + filePath);
        }

        byte[] bytes = null;
        InputStream is = null;
        try {
            is = new FileInputStream( filePath);
            bytes = IOUtils.toByteArray( is);
        } catch ( Exception e) {
            throw new ParseException( cell, e);
        } finally {
            try {
                is.close();
            } catch ( IOException e) {
                throw new ParseException( cell, e);
            }
        }

        int pictureIdx = workbook.addPicture( bytes, format);

        CreationHelper helper = workbook.getCreationHelper();

        @SuppressWarnings( "rawtypes")
        Drawing drawing = drawingCash.get( sheet);
        if ( drawing == null) {
            drawing = sheet.createDrawingPatriarch();
            drawingCash.put( sheet, drawing);
        }

        ClientAnchor anchor = helper.createClientAnchor();

        Optional<CellRangeAddress> mergedRegionIncludesTargetCell = sheet.getMergedRegions().stream()
            .filter( c -> c.isInRange( cell))
            .findFirst();
        if ( mergedRegionIncludesTargetCell.isPresent()) {
            CellRangeAddress region = mergedRegionIncludesTargetCell.get();
            anchor.setRow1( region.getFirstRow());
            anchor.setCol1( region.getFirstColumn());
            anchor.setRow2( region.getLastRow() + 1);
            anchor.setCol2( region.getLastColumn() + 1);
        }
        else {
            anchor.setRow1( cell.getRowIndex());
            anchor.setCol1( cell.getColumnIndex());
            anchor.setRow2( cell.getRowIndex() + 1);
            anchor.setCol2( cell.getColumnIndex() + 1);
        }

        if ( dx1 != null) {
            anchor.setDx1( dx1);
        }
        if ( dy1 != null) {
            anchor.setDy1( dy1);
        }

        Picture picture = drawing.createPicture( anchor, pictureIdx);
        if ( resizeBase == ResizeBase.IMAGE) {
            picture.resize();
        }
        if (scale != 1.0) {
        	picture.resize( scale);
        }
    }

    /*
     * (non-Javadoc)
     * 
     * @see org.bbreak.excella.reports.tag.ReportsTagParser#useControlRow()
     */
    @Override
    public boolean useControlRow() {
        return false;
    }

}
