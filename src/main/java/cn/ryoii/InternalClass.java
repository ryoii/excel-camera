package cn.ryoii;


import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.awt.Color;
import java.awt.Font;
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
    private int align;

    private HSSFColor ftColor;
    private HSSFColor bgColor;
    private HSSFFont font;
    private double fontZoom;
    private String text;

    public Color getAwtBgColor() {
        return bgColor == null ? null : getAwtColor(bgColor);
    }

    public Color getAwtFtColor() {
        return ftColor == null ? null : getAwtColor(ftColor);
    }

    public Font getAwtFnt() {
        if (font == null) {
            return null;
        }
        return new Font(
                this.font.getFontName(),
                Font.PLAIN,
                (int) (this.fontZoom * this.font.getFontHeightInPoints())
        );
    }

    private Color getAwtColor(HSSFColor color) {
        short[] triplet = color.getTriplet();
        return new Color(triplet[0], triplet[1], triplet[2]);
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
            bgColor = (HSSFColor) cs.getFillForegroundColorColor();
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

        org.apache.poi.ss.usermodel.Font font = workbook.getFontAt(cs.getFontIndex());
        this.font = (HSSFFont) font;
        ftColor = this.font.getHSSFColor((HSSFWorkbook) cell.getSheet().getWorkbook());
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

        switch (cell.getCellStyle().getAlignment()) {
            case RIGHT:
                align = 1;
                break;
            case CENTER:
                align = 0;
                break;
            default:
                align = -1;
        }

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

    public int getAlign() {
        return align;
    }

    public void setAlign(int align) {
        this.align = align;
    }

    public HSSFColor getFtColor() {
        return ftColor;
    }

    public void setFtColor(HSSFColor ftColor) {
        this.ftColor = ftColor;
    }

    public HSSFColor getBgColor() {
        return bgColor;
    }

    public void setBgColor(HSSFColor bgColor) {
        this.bgColor = bgColor;
    }

    public HSSFFont getFont() {
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
