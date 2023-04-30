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

package org.bbreak.excella.reports;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.function.Function;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.bbreak.excella.reports.processor.CheckMessage;
import org.bbreak.excella.reports.processor.ReportsCheckException;

/**
 * テスト用のユーティリティクラス
 * 
 * @since 1.0
 */
public class ReportsTestUtil {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( ReportsTestUtil.class);

    /**
     * XSSF最大列数
     */
    public static final int XSSF_MAX_COLUMN_NUMBER = 16384; // 2^14

    /**
     * HSSF最大列数
     */
    public static final int HSSF_MAX_COLUMN_NUMBER = 256; // 2^8

    /**
     * シート検査
     * 
     * @param expected 期待値シート
     * @param actual 実測値シート
     * @param isActCopyOfExp 実測値シートが同じワークブック上の期待値シートのコピーならtrue
     * @throws ReportsCheckException 検査エラー
     */
    public static void checkSheet( Sheet expected, Sheet actual, boolean isActCopyOfExp) throws ReportsCheckException {

        List<CheckMessage> errors = new ArrayList<CheckMessage>();

        Workbook expectedWorkbook = expected.getWorkbook();
        Workbook actualWorkbook = actual.getWorkbook();

        if ( log.isDebugEnabled()) {
            log.debug( "シート[" + actualWorkbook.getSheetName( actualWorkbook.getSheetIndex( actual)) + "] check start!");
        }

        // ----------------------
        // シート単位のチェック
        // ----------------------
        // シート名
        String eSheetName = expectedWorkbook.getSheetName( expectedWorkbook.getSheetIndex( expected));
        String aSheetName = actualWorkbook.getSheetName( actualWorkbook.getSheetIndex( actual));

        if ( !isActCopyOfExp) {
            if ( !eSheetName.equals( aSheetName)) {
                errors.add( new CheckMessage( "シート名", eSheetName, aSheetName));
            }
        }

        // 印刷設定
        String ePrintSetupString = getPrintSetupString( expected.getPrintSetup());
        String aPrintSetupString = getPrintSetupString( actual.getPrintSetup());

        if ( !ePrintSetupString.equals( aPrintSetupString)) {
            errors.add( new CheckMessage( "印刷設定", ePrintSetupString, aPrintSetupString));
        }

        // ヘッダ、フッタ
        String eHeaderString = getHeaderString( expected.getHeader());
        String aHeaderString = getHeaderString( actual.getHeader());
        if ( !eHeaderString.equals( aHeaderString)) {
            errors.add( new CheckMessage( "ヘッダ", eHeaderString, aHeaderString));
        }
        String eFooterString = getFooterString( expected.getFooter());
        String aFooterString = getFooterString( actual.getFooter());
        if ( !eFooterString.equals( aFooterString)) {
            errors.add( new CheckMessage( "フッタ", eFooterString, aFooterString));
        }

        // 改ページ
        String eBreaksString = getBreaksString( expected);
        String aBreaksString = getBreaksString( actual);
        log.debug( eBreaksString + "/" + aBreaksString);
        if ( !eBreaksString.equals( aBreaksString)) {
            errors.add( new CheckMessage( "改ページ", eBreaksString, aBreaksString));
        }

        // 印刷範囲
        String expectedPrintArea = expectedWorkbook.getPrintArea( expectedWorkbook.getSheetIndex( expected));
        String actualPrintArea = actualWorkbook.getPrintArea( actualWorkbook.getSheetIndex( actual));
        if ( expectedPrintArea != null || actualPrintArea != null) {
            // 実測値シートが期待値シートのコピーの場合、実測値シートの印刷範囲がNullになるバグのため修正されるまでチェックしない
            // if ( expectedPrintArea == null || actualPrintArea == null || !equalPrintArea( expectedPrintArea, actualPrintArea, isActCopyOfExp)) {
            // errors.add( new CheckMessage( "印刷範囲", expectedPrintArea, actualPrintArea));
            // }
            if ( !isActCopyOfExp) {
                if ( expectedPrintArea == null || actualPrintArea == null || !expectedPrintArea.equals( actualPrintArea)) {
                    errors.add( new CheckMessage( "印刷範囲", expectedPrintArea, actualPrintArea));
                }
            }
        }

        // ウィンドウ枠(固定、分割)
        String ePaneInformationString = getPaneInformationString( expected.getPaneInformation());
        String aPaneInformationString = getPaneInformationString( actual.getPaneInformation());

        if ( !ePaneInformationString.equals( aPaneInformationString)) {
            errors.add( new CheckMessage( "ウィンドウ枠(固定、分割)", expectedPrintArea, actualPrintArea));
        }

        // 行タイトル、列タイトル・・・セットはできるが。。。

        // 表示ズーム・・・セットはできるが。。。

        // アウトライン・・・セットはできるが。。。

        // セルコメント

        // グリッドライン表示
        if ( expected.isDisplayGridlines() ^ actual.isDisplayGridlines()) {
            errors.add( new CheckMessage( "グリッドライン表示", String.valueOf( expected.isDisplayGridlines()), String.valueOf( actual.isDisplayGridlines())));
        }

        // 見出し表示
        if ( expected.isDisplayRowColHeadings() ^ actual.isDisplayRowColHeadings()) {
            errors.add( new CheckMessage( "見出し表示", String.valueOf( expected.isDisplayRowColHeadings()), String.valueOf( actual.isDisplayRowColHeadings())));
        }

        // 数式表示
        if ( expected.isDisplayFormulas() ^ actual.isDisplayFormulas()) {
            errors.add( new CheckMessage( "数式表示", String.valueOf( expected.isDisplayFormulas()), String.valueOf( actual.isDisplayFormulas())));
        }
        // 結合セル
        if ( expected.getNumMergedRegions() != actual.getNumMergedRegions()) {
            errors.add( new CheckMessage( "結合セル数", String.valueOf( expected.getNumMergedRegions()), String.valueOf( actual.getNumMergedRegions())));
        }

        for ( int i = 0; i < actual.getNumMergedRegions(); i++) {

            CellRangeAddress actualAddress = actual.getMergedRegion(i);

            StringBuffer expectedAdressBuffer = new StringBuffer();
            boolean equalAddress = false;
            for ( int j = 0; j < expected.getNumMergedRegions(); j++) {
                CellRangeAddress expectedAddress = expected.getMergedRegion(j);

                if ( expectedAddress.toString().equals( actualAddress.toString())) {
                    equalAddress = true;
                    break;
                }
                CellReference crA = new CellReference( expectedAddress.getFirstRow(), expectedAddress.getFirstColumn());
                CellReference crB = new CellReference( expectedAddress.getLastRow(), expectedAddress.getLastColumn());
                expectedAdressBuffer.append( " [" + crA.formatAsString() + ":" + crB.formatAsString() + "]");
            }

            if ( !equalAddress) {
                errors.add( new CheckMessage( "結合セル", expectedAdressBuffer.toString(), actualAddress.toString()));
            }

        }

        int maxColumnNum = -1;
        if ( expected instanceof HSSFSheet) {
            maxColumnNum = HSSF_MAX_COLUMN_NUMBER;
        } else if ( expected instanceof XSSFSheet) {
            maxColumnNum = XSSF_MAX_COLUMN_NUMBER;
        }
        for ( int i = 0; i < maxColumnNum; i++) {
            try {
                // 列スタイル
                checkCellStyle( expected.getWorkbook(), expected.getColumnStyle( i), actual.getWorkbook(), actual.getColumnStyle( i));
            } catch ( ReportsCheckException e) {
                CheckMessage checkMessage = e.getCheckMessages().iterator().next();
                checkMessage.setMessage( "列[" + i + "]" + checkMessage.getMessage());
                errors.add( checkMessage);
            }

            // 列幅
            if ( expected.getColumnWidth( i) != actual.getColumnWidth( i)) {
                errors.add( new CheckMessage( "列幅[" + i + "]", String.valueOf( expected.getColumnWidth( i)), String.valueOf( actual.getColumnWidth( i))));
            }
        }

        // 行単位チェック
        if ( expected.getLastRowNum() != actual.getLastRowNum()) {
            // 期待値まで、値がブランクか調査する
            if ( expected.getLastRowNum() < actual.getLastRowNum()) {
                int lastRowIndex = -1;
                if ( expected instanceof HSSFSheet) {
                    lastRowIndex = 0;
                }
                Iterator<Row> rowIterator = actual.rowIterator();
                while ( rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    // すべてブランクは、空白行
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while ( cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        if ( cell.getCellType() != CellType.BLANK) {
                            lastRowIndex = row.getRowNum();
                            break;
                        }
                    }
                }
                if ( expected.getLastRowNum() != lastRowIndex) {
                    errors.add( new CheckMessage( "最終行", String.valueOf( expected.getLastRowNum()), String.valueOf( lastRowIndex)));
                }
            } else {
                errors.add( new CheckMessage( "最終行", String.valueOf( expected.getLastRowNum()), String.valueOf( actual.getLastRowNum())));
            }

        }

        if ( errors.isEmpty()) {
            for ( int i = 0; i <= expected.getLastRowNum(); i++) {
                try {
                    checkRow( expected.getRow( i), actual.getRow( i));
                } catch ( ReportsCheckException e) {
                    errors.addAll( e.getCheckMessages());
                }
            }
        }

        if ( !errors.isEmpty()) {
            if ( log.isErrorEnabled()) {
                for ( CheckMessage message : errors) {
                    log.error( "相違有[" + message.getMessage() + "]");
                    log.error( "期待値:" + message.getExpected());
                    log.error( "実測値:" + message.getActual());
                }
            }
            throw new ReportsCheckException( errors);
        }

        if ( log.isDebugEnabled()) {
            log.debug( "シート[" + actualWorkbook.getSheetName( actualWorkbook.getSheetIndex( actual)) + "] check end.");
        }

    }

    /**
     * 行検査
     * 
     * @param expected 期待値行
     * @param actual 実測値行
     * @throws ReportsCheckException 検査エラー
     */
    public static void checkRow( Row expected, Row actual) throws ReportsCheckException {

        List<CheckMessage> errors = new ArrayList<CheckMessage>();

        // ----------------------
        // 行単位のチェック
        // ----------------------

        if ( expected == null && actual == null) {
            return;
        }

        if ( expected == null) {
            if ( actual.iterator().hasNext()) {
                errors.add( new CheckMessage( "行[" + actual.getRowNum() + "]", null, actual.toString()));
                throw new ReportsCheckException( errors);
            } else {
                return;
            }
        }
        if ( actual == null) {
            if ( expected.iterator().hasNext()) {
                errors.add( new CheckMessage( "行[" + expected.getRowNum() + "]", expected.toString(), null));
                throw new ReportsCheckException( errors);
            } else {
                return;
            }
        }

        // 行の高さ(shiftRowすると行の高さがおかしくなるため、チェックしない)
        // float adjustHight = 0f;
        // if ( hasHeightAdjustBorderCell( actual.getSheet().getRow( actual.getRowNum() - 1), actual, actual.getSheet().getRow( actual.getRowNum() + 1))) {
        // log.error( "true");
        // adjustHight = 0.75f;
        // }
        //
        // if ( expected.getHeightInPoints() != actual.getHeightInPoints() + adjustHight) {
        // if ( log.isErrorEnabled()) {
        // log.error( "expectedROW[" + expected.getRowNum() + "]:" + expected.getHeightInPoints());
        // log.error( "actualROW[" + actual.getRowNum() + "]:" + (actual.getHeightInPoints() + adjustHight));
        // }
        // throw new Exception( "行の高さ");
        // }

        // 最終セル
        if ( expected.getLastCellNum() != actual.getLastCellNum()) {
            errors.add( new CheckMessage( "行[" + expected.getRowNum() + "]最終セル", String.valueOf( expected.getLastCellNum()), String.valueOf( actual.getLastCellNum())));
            throw new ReportsCheckException( errors);
        }

        // セル単位チェック
        for ( int i = 0; i < expected.getLastCellNum(); i++) {
            try {
                checkCell( expected.getCell( i), actual.getCell( i));
            } catch ( ReportsCheckException e) {
                errors.addAll( e.getCheckMessages());
            }
        }
        if ( !errors.isEmpty()) {
            throw new ReportsCheckException( errors);
        }
    }

    /**
     * セル検査
     * 
     * @param expected 期待値セル
     * @param actual 実測値セル
     * @throws ReportsCheckException 検査エラー
     */
    public static void checkCell( Cell expected, Cell actual) throws ReportsCheckException {

        List<CheckMessage> errors = new ArrayList<CheckMessage>();

        // ----------------------
        // セル単位のチェック
        // ----------------------

        if ( expected == null && actual == null) {
            return;
        }

        if ( expected == null) {
            // if(actual.getCellStyle() != null || actual.getCellType() != Cell.CELL_TYPE_BLANK){
            errors.add( new CheckMessage( "セル(" + actual.getRowIndex() + "," + actual.getColumnIndex() + ")", null, actual.toString()));
            throw new ReportsCheckException( errors);
            // }
        }
        if ( actual == null) {
            errors.add( new CheckMessage( "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")", expected.toString(), null));
            throw new ReportsCheckException( errors);
        }

        // 型
        if ( expected.getCellType() != actual.getCellType()) {
            errors.add( new CheckMessage( "型[" + "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + "]", getCellTypeString( expected.getCellType()),
                getCellTypeString( actual.getCellType())));
            throw new ReportsCheckException( errors);
        }

        try {
            checkCellStyle( expected.getRow().getSheet().getWorkbook(), expected.getCellStyle(), actual.getRow().getSheet().getWorkbook(), actual.getCellStyle());
        } catch ( ReportsCheckException e) {
            CheckMessage checkMessage = e.getCheckMessages().iterator().next();
            checkMessage.setMessage( "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + checkMessage.getMessage());
            errors.add( checkMessage);
            throw new ReportsCheckException( errors);
        }

        // 値
        if ( !getCellValue( expected).equals( getCellValue( actual))) {
            log.error( getCellValue( expected) + " / " + getCellValue( actual));
            errors.add( new CheckMessage( "値[" + "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + "]", String.valueOf( getCellValue( expected)), String
                .valueOf( getCellValue( actual))));
            throw new ReportsCheckException( errors);
        }

    }

    /**
     * スタイル検査
     * 
     * @param expectedWorkbook 期待値ブック
     * @param expected 期待値スタイル
     * @param actualWorkbook 実測値スタイルブック
     * @param actual 実測値スタイル
     * @throws ReportsCheckException 検査エラー
     */
    private static void checkCellStyle( Workbook expectedWorkbook, CellStyle expected, Workbook actualWorkbook, CellStyle actual) throws ReportsCheckException {

        List<CheckMessage> errors = new ArrayList<CheckMessage>();

        if ( expected == null && actual == null) {
            return;
        }

        if ( expected == null) {
            errors.add( new CheckMessage( "スタイル", null, actual.toString()));
            throw new ReportsCheckException( errors);
        }

        if ( actual == null) {
            errors.add( new CheckMessage( "スタイル", expected.toString(), null));
            throw new ReportsCheckException( errors);
        }

        String eCellStyleString = getCellStyleString(expectedWorkbook, expected);
        String aCellStyleString = getCellStyleString(actualWorkbook, actual);

        if ( !eCellStyleString.equals( aCellStyleString)) {
            errors.add( new CheckMessage( "スタイル", eCellStyleString, aCellStyleString));
            throw new ReportsCheckException( errors);

        }
    }

    /**
     * スタイルの文字列表現を取得する
     * 
     * @param workbook  ブック
     * @param cellStyle スタイル
     * @return スタイルの文字列表現
     */
    private static String getCellStyleString(Workbook workbook, CellStyle cellStyle) {
        StringBuffer sb = new StringBuffer();
        if (cellStyle != null) {
            sb.append("Font=").append(getFontString(workbook, cellStyle)).append(",");
            sb.append("DataFormat=").append(cellStyle.getDataFormat()).append(",");
            sb.append("DataFormatString=").append(cellStyle.getDataFormatString()).append(",");
            sb.append("Hidden=").append(cellStyle.getHidden()).append(",");
            sb.append("Locked=").append(cellStyle.getLocked()).append(",");
            sb.append("Alignment=").append(cellStyle.getAlignment()).append(",");
            sb.append("WrapText=").append(cellStyle.getWrapText()).append(",");
            sb.append("VerticalAlignment=").append(cellStyle.getVerticalAlignment()).append(",");
            sb.append("Rotation=").append(cellStyle.getRotation()).append(",");
            sb.append("Indention=").append(cellStyle.getIndention()).append(",");
            sb.append("BorderLeft=").append(cellStyle.getBorderLeft()).append(",");
            sb.append("BorderRight=").append(cellStyle.getBorderRight()).append(",");
            sb.append("BorderTop=").append(cellStyle.getBorderTop()).append(",");
            sb.append("BorderBottom=").append(cellStyle.getBorderBottom()).append(",");

            sb.append("LeftBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.LEFT_BORDER))
                    .append(",");
            sb.append("RightBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.RIGHT_BORDER))
                    .append(",");
            sb.append("TopBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.TOP_BORDER))
                    .append(",");
            sb.append("BottomBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.BOTTOM_BORDER))
                    .append(",");

            sb.append("FillPattern=").append(cellStyle.getFillPattern()).append(",");
            sb.append("FillForegroundColor=").append(getColorString(workbook, cellStyle, ColorPosition.FILL_FOREGROUND))
                    .append(",");
            sb.append("FillBackgroundColor=")
                    .append(getColorString(workbook, cellStyle, ColorPosition.FILL_BACKGROUND));
        }
        return sb.toString();
    }

    private enum ColorPosition {
        LEFT_BORDER(CellStyle::getLeftBorderColor, XSSFCellStyle::getLeftBorderXSSFColor),
        RIGHT_BORDER(CellStyle::getRightBorderColor, XSSFCellStyle::getRightBorderXSSFColor),
        TOP_BORDER(CellStyle::getTopBorderColor, XSSFCellStyle::getTopBorderXSSFColor),
        BOTTOM_BORDER(CellStyle::getBottomBorderColor, XSSFCellStyle::getBottomBorderXSSFColor),
        FILL_FOREGROUND(CellStyle::getFillForegroundColor, XSSFCellStyle::getFillForegroundXSSFColor),
        FILL_BACKGROUND(CellStyle::getFillBackgroundColor, XSSFCellStyle::getFillBackgroundXSSFColor);

        private final Function<CellStyle, Short> colorIndexAccessor;
        private final Function<XSSFCellStyle, XSSFColor> xssfColorAccessor;

        ColorPosition(Function<CellStyle, Short> colorIndexAccessor,
                Function<XSSFCellStyle, XSSFColor> xssfColorAccessor) {
            this.colorIndexAccessor = colorIndexAccessor;
            this.xssfColorAccessor = xssfColorAccessor;
        }

        private short getColorIndex(CellStyle cellStyle) {
            return colorIndexAccessor.apply(cellStyle);
        }

        private XSSFColor getXSSFColor(CellStyle cellStyle) {
            if (cellStyle instanceof XSSFCellStyle) {
                return xssfColorAccessor.apply((XSSFCellStyle) cellStyle);
            }
            throw new IllegalArgumentException("cellStyle is not instanceof XSSFCellStyle: " + cellStyle.getClass());
        }
    }

    private static String getFontString(Workbook workbook, CellStyle cellStyle) {
        if (cellStyle instanceof HSSFCellStyle) {
            HSSFFont font = ((HSSFCellStyle) cellStyle).getFont(workbook);
            return getHSSFFontString((HSSFWorkbook) workbook, font);
        } else if (cellStyle instanceof XSSFCellStyle) {
            XSSFFont font = ((XSSFCellStyle) cellStyle).getFont();
            return font.getCTFont().toString();
        }
        return "";
    }

    /**
     * HSSFフォントの文字列表現を取得する
     * 
     * @param workbook ブック
     * @param font     フォント
     * @return フォントの文字列表現
     */
    private static String getHSSFFontString(HSSFWorkbook workbook, HSSFFont font) {
        StringBuffer sb = new StringBuffer();
        sb.append("[FONT]");
        sb.append("fontheight=").append(Integer.toHexString(font.getFontHeight())).append(",");
        sb.append("italic=").append(font.getItalic()).append(",");
        sb.append("strikout=").append(font.getStrikeout()).append(",");
        sb.append("colorpalette=").append(getHSSFColorString((HSSFWorkbook) workbook, font.getColor())).append(",");
        sb.append("boldweight=").append(font.getBold()).append(",");
        sb.append("supersubscript=").append(Integer.toHexString(font.getTypeOffset())).append(",");
        sb.append("underline=").append(Integer.toHexString(font.getUnderline())).append(",");
        sb.append("charset=").append(Integer.toHexString(font.getCharSet())).append(",");
        sb.append("fontname=").append(font.getFontName());
        sb.append("[/FONT]");
        return sb.toString();
    }

    private static String getColorString(Workbook workbook, CellStyle cellStyle, ColorPosition position) {
        if (cellStyle instanceof HSSFCellStyle) {
            return getHSSFColorString((HSSFWorkbook) workbook, position.getColorIndex(cellStyle));
        } else {
            return getXSSFColorString(position.getXSSFColor(cellStyle));
        }
    }

    /**
     * HSSF色の文字列表現を取得する
     * 
     * @param workbook ブック
     * @param index    色インデックス
     * @return HSSF色の文字列表現
     */
    private static String getHSSFColorString(HSSFWorkbook workbook, short index) {
        HSSFPalette palette = workbook.getCustomPalette();
        if (palette.getColor(index) != null) {
            HSSFColor color = palette.getColor(index);
            return color.getHexString();
        } else {
            return "";
        }
    }

    /**
     * XSSF色の文字列表現を取得する
     * 
     * @param color XSSF色
     * @return XSSF色の文字列表現
     */
    private static String getXSSFColorString(XSSFColor color) {
        StringBuffer sb = new StringBuffer("[");
        if (color != null) {
            sb.append("Indexed=").append(color.getIndexed()).append(",");
            sb.append("Rgb=");
            if (color.getRGB() != null) {
                for (byte b : color.getRGB()) {
                    sb.append(String.format("%02x", b).toUpperCase());
                }
            }
            sb.append(",");
            sb.append("Tint=").append(color.getTint()).append(",");
            sb.append("Theme=").append(color.getTheme()).append(",");
            sb.append("Auto=").append(color.isAuto());
        }
        return sb.append("]").toString();
    }

    /**
     * 印刷設定の文字列表現を取得する。
     * 
     * @param printSetup 印刷設定
     * @return 印刷設定の文字列表現
     */
    public static String getPrintSetupString( PrintSetup printSetup) {
        StringBuffer sb = new StringBuffer();
        if ( printSetup != null) {
            sb.append( "PaperSize=").append( printSetup.getPaperSize()).append( ",");
            sb.append( "Scale=").append( printSetup.getScale()).append( ",");
            sb.append( "PageStart=").append( printSetup.getPageStart()).append( ",");
            sb.append( "FitWidth=").append( printSetup.getFitWidth()).append( ",");
            sb.append( "FitHeight=").append( printSetup.getFitHeight()).append( ",");
            sb.append( "LeftToRight=").append( printSetup.getLeftToRight()).append( ",");
            sb.append( "Landscape=").append( printSetup.getLandscape()).append( ",");
            sb.append( "ValidSettings=").append( printSetup.getValidSettings()).append( ",");
            sb.append( "NoColor=").append( printSetup.getNoColor()).append( ",");
            sb.append( "Draft=").append( printSetup.getDraft()).append( ",");
            sb.append( "Notes=").append( printSetup.getNotes()).append( ",");
            sb.append( "NoOrientation=").append( printSetup.getNoOrientation()).append( ",");
            sb.append( "UsePage=").append( printSetup.getUsePage()).append( ",");
            sb.append( "HResolution=").append( printSetup.getHResolution()).append( ",");
            sb.append( "VResolution=").append( printSetup.getVResolution()).append( ",");
            sb.append( "HeaderMargin=").append( printSetup.getHeaderMargin()).append( ",");
            sb.append( "FooterMargin=").append( printSetup.getFooterMargin()).append( ",");
            sb.append( "Copies=").append( printSetup.getCopies());
        }
        return sb.toString();

    }

    /**
     * ヘッダ情報の文字列表現を取得する。
     * 
     * @param header ヘッダ情報
     * @return ヘッダ情報の文字列表現
     */
    private static String getHeaderString( Header header) {
        if ( header != null) {
            return "left=" + header.getLeft() + ",center=" + header.getCenter() + ",right=" + header.getRight();
        } else {
            return "";
        }
    }

    /**
     * @param footer
     * @return
     */
    private static String getFooterString( Footer footer) {
        if ( footer != null) {
            return "left=" + footer.getLeft() + ",center=" + footer.getCenter() + ",right=" + footer.getRight();
        } else {
            return "";
        }
    }

    /**
     * 改ページ情報の文字列表現を取得する。
     * 
     * @param sheet シート
     * @return 改ページ情報の文字列表現
     */
    private static String getBreaksString( Sheet sheet) {

        return "row=" + Arrays.toString( sheet.getRowBreaks()) + ",column=" + Arrays.toString( sheet.getColumnBreaks());

    }

    /**
     * ウィンドウ枠の文字列表現を取得する。
     * 
     * @param information ウィンドウ枠
     * @return ウィンドウ枠の文字列表現
     */
    private static String getPaneInformationString( PaneInformation information) {
        StringBuffer sb = new StringBuffer();
        if ( information != null) {
            sb.append( "HorizontalSplitPosition=").append( information.getHorizontalSplitPosition()).append( ",");
            sb.append( "HorizontalSplitTopRow=").append( information.getHorizontalSplitTopRow()).append( ",");
            sb.append( "VerticalSplitPosition=").append( information.getVerticalSplitPosition()).append( ",");
            sb.append( "VerticalSplitLeftColumn=").append( information.getVerticalSplitLeftColumn());
        }
        return sb.toString();
    }

    /**
     * セルの値の文字列表現を取得する
     * 
     * @param cell セル
     * @return セルの値の文字列表現
     */
    private static String getCellValue( Cell cell) {
        String value = null;

        if ( cell != null) {
            switch ( cell.getCellType()) {
                case BLANK:
                    value = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    value = String.valueOf( cell.getBooleanCellValue());
                    break;
                case ERROR:
                    value = String.valueOf( cell.getErrorCellValue());
                    break;
                case NUMERIC:
                    value = String.valueOf( cell.getNumericCellValue());
                    break;
                case STRING:
                    value = cell.getStringCellValue();
                    break;
                case FORMULA:
                    value = cell.getCellFormula();
                default:
                    value = "";
            }
        }
        return value;
    }

    /**
     * セルタイプの文字列表現を取得する
     * 
     * @param cellType セルタイプ
     * @return セルタイプの文字列表現
     */
    private static String getCellTypeString( CellType cellType) {
        if ( cellType != null) {
            return cellType.toString();
        } else {
            return "";
        }
    }

    public static String getTestOutputDir() {

        String tempDir = System.getProperty( "user.dir") + "/work/test/";
        File file = new File( tempDir);
        if ( !file.exists()) {
            file.mkdirs();
        }

        return tempDir;
    }

    /**
     * 印刷範囲検査
     * 
     * @param expectedPrintArea 期待値印刷範囲
     * @param actualPrintArea 実測値印刷範囲
     * @param isCopy 実測値シートが期待値シートのコピーならtrue
     * @return
     */
    // private static boolean equalPrintArea( String expectedPrintArea, String actualPrintArea, boolean isActCopyOfExp) {
    // boolean returnValue = true;
    // if ( isActCopyOfExp) {
    // String expPrintAreaWithoutSheetName = expectedPrintArea.substring( expectedPrintArea.indexOf( '!'));
    // String actPrintAreaWithoutSheetName = actualPrintArea.substring( actualPrintArea.indexOf( '!'));
    // if ( !expPrintAreaWithoutSheetName.equals( actPrintAreaWithoutSheetName)) {
    // returnValue = false;
    // }
    // } else {
    // if ( !expectedPrintArea.equals( actualPrintArea)) {
    // returnValue = false;
    // }
    // }
    // return returnValue;
    // }
}
