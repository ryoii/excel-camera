/*
 * Copyright (c) 2021 ryoii
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

package cn.ryoii;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.text.DecimalFormat;

/**
 * @author ryoii
 */
class Grid {

    private Workbook workbook;

    private int x;
    private int y;
    private int width;
    private int height;
    private int row;
    private int col;
    private int border;
    private HorizontalAlignment horizontalAlign;
    private VerticalAlignment verticalAlignment;

    private org.apache.poi.ss.usermodel.Color ftColor;
    private org.apache.poi.ss.usermodel.Color bgColor;
    private org.apache.poi.ss.usermodel.Font font;
    private double fontZoom;
    private String text;

    public java.awt.Color getAwtBgColor() {
        return bgColor == null ? null : getAwtColor(bgColor);
    }

    public java.awt.Color getAwtFtColor() {
        return ftColor == null ? null : getAwtColor(ftColor);
    }

    public java.awt.Font getAwtFnt() {
        if (font == null) {
            return null;
        }
        return new java.awt.Font(
                this.font.getFontName(),
                java.awt.Font.PLAIN,
                (int) (this.fontZoom * this.font.getFontHeightInPoints())
        );
    }

    private java.awt.Color getAwtColor(Color color) {
        int[] rgb = getRgb(color);
        return new java.awt.Color(rgb[0], rgb[1], rgb[2]);
    }

    private int[] getRgb(org.apache.poi.ss.usermodel.Color color) {
        int[] res = new int[]{0, 0, 0};
        if (color instanceof HSSFColor) {
            short[] rgb = ((HSSFColor) color).getTriplet();
            res = new int[]{rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF};
        } else if (color instanceof XSSFColor) {
            byte[] rgb = ((XSSFColor) color).getRGBWithTint();
            res = new int[]{rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF};
        }
        return res;
    }

    public static Grid build(Workbook workbook, UserCell cell) {
        Grid grid = new Grid();
        grid.setWorkbook(workbook);

        grid.setX(cell.getLeft());
        grid.setY(cell.getTop());
        grid.setWidth(cell.getRight() - cell.getLeft());
        grid.setHeight(cell.getBottom() - cell.getTop());
        grid.setRow(cell.getRow());
        grid.setCol(cell.getCol());

        if (cell.getCell() != null) {
            grid.backgroundColor(cell.getCell());
            grid.border(cell.getCell());
            grid.font(cell.getCell());
            grid.text(cell.getCell());
        }

        return grid;
    }

    private void backgroundColor(Cell cell) {
        CellStyle cs = cell.getCellStyle();
        if (cs.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
            bgColor = cs.getFillForegroundColorColor();
        }
    }

    private void border(Cell cell) {
        CellStyle cs = cell.getCellStyle();
        if (cs.getBorderTop() != BorderStyle.NONE) {
            border |= 0x1;
        }
        if (cs.getBorderRight() != BorderStyle.NONE) {
            border |= 0x2;
        }
        if (cs.getBorderBottom() != BorderStyle.NONE) {
            border |= 0x4;
        }
        if (cs.getBorderLeft() != BorderStyle.NONE) {
            border |= 0x8;
        }
    }

    private void font(Cell cell) {
        CellStyle cs = cell.getCellStyle();

        this.font = workbook.getFontAt(cs.getFontIndex());
        if (font instanceof HSSFFont) {
            ftColor = ((HSSFFont) this.font).getHSSFColor((HSSFWorkbook) cell.getSheet().getWorkbook());
        } else if (font instanceof XSSFFont) {
            ftColor = ((XSSFFont) this.font).getXSSFColor();
        }

    }

    private void text(Cell cell) {
        String strCell;
        switch (cell.getCellType()) {
            case NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue());
                break;
            case STRING:
                strCell = cell.getStringCellValue();
                break;
            case BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                try {
                    strCell = String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    strCell = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            default:
                strCell = "";
        }

        horizontalAlign = cell.getCellStyle().getAlignment();
        verticalAlignment = cell.getCellStyle().getVerticalAlignment();

        if (cell.getCellStyle().getDataFormatString().contains("0.00%")) {
            try {
                double dbCell = Double.parseDouble(strCell);
                strCell = new DecimalFormat("#.00").format(dbCell * 100) + "%";
            } catch (NumberFormatException ignored) {
            }
        }

        this.text = strCell.matches("\\w*\\.0") ? strCell.substring(0, strCell.length() - 2) : strCell;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public int getX() {
        return x;
    }

    public void setX(int x) {
        this.x = x;
    }

    public int getY() {
        return y;
    }

    public void setY(int y) {
        this.y = y;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public int getBorder() {
        return border;
    }

    public void setBorder(int border) {
        this.border = border;
    }

    public HorizontalAlignment getHorizontalAlign() {
        return horizontalAlign;
    }

    public Grid setHorizontalAlign(HorizontalAlignment horizontalAlign) {
        this.horizontalAlign = horizontalAlign;
        return this;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public Grid setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    public Color getFtColor() {
        return ftColor;
    }

    public void setFtColor(HSSFColor ftColor) {
        this.ftColor = ftColor;
    }

    public Color getBgColor() {
        return bgColor;
    }

    public void setBgColor(HSSFColor bgColor) {
        this.bgColor = bgColor;
    }

    public org.apache.poi.ss.usermodel.Font getFont() {
        return font;
    }

    public void setFont(HSSFFont font) {
        this.font = font;
    }

    public double getFontZoom() {
        return fontZoom;
    }

    public Grid setFontZoom(double fontZoom) {
        this.fontZoom = fontZoom;
        return this;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public boolean hashBorderTop() {
        return (border & 0x01) != 0;
    }

    public void setBorderTop() {
        border |= 0x01;
    }

    public boolean hashBorderRight() {
        return (border & 0x02) != 0;
    }

    public void setBorderRight() {
        border |= 0x02;
    }

    public boolean hashBorderBottom() {
        return (border & 0x04) != 0;
    }

    public void setBorderBottom() {
        border |= 0x04;
    }

    public boolean hashBorderLeft() {
        return (border & 0x08) != 0;
    }

    public void setBorderLeft() {
        border |= 0x08;
    }
}

class UserCell {

    private int row;
    private int col;

    private int top;
    private int right;
    private int bottom;
    private int left;

    private Cell cell;

    public static UserCell build(Sheet sheet, int row, int col) {
        UserCell cell = new UserCell();
        if (sheet.getRow(row) == null) {
            cell.setCell(null);
        } else {
            cell.setCell(sheet.getRow(row).getCell(col));
        }
        cell.setRow(row);
        cell.setCol(col);
        return cell;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public int getTop() {
        return top;
    }

    public void setTop(int top) {
        this.top = top;
    }

    public int getRight() {
        return right;
    }

    public void setRight(int right) {
        this.right = right;
    }

    public int getBottom() {
        return bottom;
    }

    public void setBottom(int bottom) {
        this.bottom = bottom;
    }

    public int getLeft() {
        return left;
    }

    public void setLeft(int left) {
        this.left = left;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }
}
