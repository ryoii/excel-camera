package cn.ryoii;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

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
    private boolean middle;

    private XSSFColor ftColor;
    private XSSFColor bgColor;
    private XSSFFont font;
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
        return new Font(this.font.getFontName(), Font.PLAIN, this.font.getFontHeightInPoints());
    }

    private Color getAwtColor(XSSFColor color) {
        byte[] bytes = color.getRGB();
        int rgb = 0;
        for (byte b : bytes) {
            rgb = (rgb << 8) | b;
        }
        return new Color(rgb);
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
            bgColor = (XSSFColor) cs.getFillForegroundColorColor();
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
        this.font = (XSSFFont) font;
        ftColor = this.font.getXSSFColor();
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

        this.middle = cell.getCellStyle().getAlignment() == HorizontalAlignment.CENTER;

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

    public boolean isMiddle() {
        return middle;
    }

    public void setMiddle(boolean middle) {
        this.middle = middle;
    }

    public XSSFColor getFtColor() {
        return ftColor;
    }

    public void setFtColor(XSSFColor ftColor) {
        this.ftColor = ftColor;
    }

    public XSSFColor getBgColor() {
        return bgColor;
    }

    public void setBgColor(XSSFColor bgColor) {
        this.bgColor = bgColor;
    }

    public XSSFFont getFont() {
        return font;
    }

    public void setFont(XSSFFont font) {
        this.font = font;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
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