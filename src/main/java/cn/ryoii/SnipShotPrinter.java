package cn.ryoii;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import sun.awt.SunHints;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author ryoii
 */
public class SnipShotPrinter {

    private final SnipShotConfig config;
    private List<CellRangeAddress> rangeAddress;
    private final List<Grid> grids;

    private final int rowSize;
    private final int colSize;
    private int imageWidth = 0;
    private int imageHeight = 0;

    private final UserCell[][] cells;

    public SnipShotPrinter(SnipShotConfig config) {
        this.config = config;
        this.rowSize = config.getRowTo() - config.getRowFrom();
        this.colSize = config.getColTo() - config.getColFrom();
        this.grids = new ArrayList<>(this.colSize * this.rowSize);

        this.cells = new UserCell[rowSize][colSize];
    }

    public InputStream toImageStream() throws Exception {
        Workbook wb = WorkbookFactory.create(new File(config.getFileName()));
        Sheet sheet = wb.getSheet(config.getSheetName());

        // merged regions
        rangeAddress = sheet.getMergedRegions();

        for (int i = 0; i < rowSize; i++) {
            // skip empty row
            Row row = sheet.getRow(i + config.getRowFrom());
            if (row == null) {
                continue;
            }

            // calculate height
            float heightPx = row.getHeightInPoints();
            imageHeight += heightPx;

            for (int j = 0; j < colSize; j++) {
                UserCell cell = UserCell.build(sheet, i + config.getRowFrom(), j + config.getColFrom());
                cells[i][j] = cell;

                // calculate width
                float widthPx = sheet.getColumnWidthInPixels(j + config.getColFrom());
                if (i == 0) {
                    imageWidth += widthPx;
                    cell.setTop(0);
                    cell.setBottom((int) (heightPx * config.getRowZoom()));
                } else {
                    int preBottom = cells[i - 1][j].getBottom();
                    cell.setTop(preBottom);
                    cell.setBottom((int) (heightPx * config.getRowZoom() + preBottom));
                }
                if (j == 0) {
                    cell.setLeft(0);
                    cell.setRight((int) (widthPx * config.getColZoom()));
                } else {
                    int preRight = cells[i][j - 1].getRight();
                    cell.setLeft(preRight);
                    cell.setRight((int) (widthPx * config.getColZoom() + preRight));
                }
            }
        }

        imageWidth = (int) (imageWidth * config.getColZoom());
        imageHeight = (int) (imageHeight * config.getRowZoom());
        wb.close();


        for (int i = 0; i < rowSize; i++) {
            for (int j = 0; j < colSize; j++) {
                Grid grid = Grid.build(wb, cells[i][j]);
                // is merged cell ?
                int[] isInMergedStatus = isInMerged(grid.getRow(), grid.getCol());

                if (isInMergedStatus[0] == 0 && isInMergedStatus[1] == 0) {
                    // is merged and isn't the first cell in merged region
                    continue;
                } else if (isInMergedStatus[0] != -1 && isInMergedStatus[1] != -1) {
                    // is merged and is the first cell, resize it
                    int lastRowPos = Math.min(isInMergedStatus[0], rowSize - 1);
                    int lastColPos = Math.min(isInMergedStatus[1], colSize - 1);

                    grid.setWidth(cells[i][lastColPos].getRight() - grid.getX());
                    grid.setHeight(cells[lastRowPos][j].getBottom() - grid.getY());
                } // else: empty cell

                grids.add(grid);
            }
        }

        return getImageStream();
    }

    private InputStream getImageStream() throws IOException {
        BufferedImage image = new BufferedImage(imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();
        // smooth font setting
        g2d.setRenderingHint(SunHints.KEY_ANTIALIASING, SunHints.VALUE_ANTIALIAS_OFF);
        g2d.setRenderingHint(SunHints.KEY_TEXT_ANTIALIASING, SunHints.VALUE_TEXT_ANTIALIAS_DEFAULT);
        g2d.setRenderingHint(SunHints.KEY_STROKE_CONTROL, SunHints.VALUE_STROKE_DEFAULT);
        g2d.setRenderingHint(SunHints.KEY_TEXT_ANTIALIAS_LCD_CONTRAST, 140);
        g2d.setRenderingHint(SunHints.KEY_FRACTIONALMETRICS, SunHints.VALUE_FRACTIONALMETRICS_OFF);
        g2d.setRenderingHint(SunHints.KEY_RENDERING, SunHints.VALUE_RENDER_DEFAULT);

        // panel background color
        g2d.setColor(java.awt.Color.white);
        g2d.fillRect(0, 0, imageWidth, imageHeight);

        // draw the grid
        for (Grid g : grids) {
            // fill cell background color
            g2d.setColor(g.getBgColor() == null ? java.awt.Color.white : g.getAwtBgColor());
            g2d.fillRect(g.getX(), g.getY(), g.getWidth(), g.getHeight());

            // draw the border
            g2d.setColor(Color.black);
            g2d.setStroke(new BasicStroke(1));

            if ((g.getBorder() & 0x01) != 0) {
                g2d.drawLine(g.getX(), g.getY() - 1, g.getX() + g.getWidth() - 1, g.getY() - 1);
            }
            if ((g.getBorder() & 0x02) != 0) {
                g2d.drawLine(g.getX() + g.getWidth() - 1, g.getY(), g.getX() + g.getWidth() - 1, g.getY() + g.getHeight() - 1);
            }
            if ((g.getBorder() & 0x04) != 0) {
                g2d.drawLine(g.getX(), g.getY() + g.getHeight() - 1, g.getX() + g.getWidth() - 1, g.getY() + g.getHeight() - 1);
            }
            if ((g.getBorder() & 0x08) != 0) {
                g2d.drawLine(g.getX() - 1, g.getY(), g.getX() - 1, g.getY() + g.getHeight() - 1);
            }

            // draw font
            g2d.setColor(g.getAwtFtColor());
            java.awt.Font font = g.getAwtFnt();
            if (font != null) {
                FontMetrics fm = g2d.getFontMetrics(font);
                // get font width
                int strWidth = fm.stringWidth(g.getText());
                g2d.setFont(font);

                // is text-align center
                int x;
                if (g.isMiddle()) {
                    x = g.getX() + (g.getWidth() - strWidth) / 2 - 1;
                } else {
                    x = g.getX() + (g.getWidth() - strWidth - 1);
                }
                g2d.drawString(g.getText(), x,
                        g.getY() + (g.getHeight() - font.getSize()) / 2 + font.getSize());
            }
        }

        g2d.dispose();

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", bos);
        return new ByteArrayInputStream(bos.toByteArray());
    }


    private int[] isInMerged(int row, int col) {
        int[] isInMergedStatus = {-1, -1};
        for (CellRangeAddress cra : rangeAddress) {
            if (row == cra.getFirstRow() && col == cra.getFirstColumn()) {
                isInMergedStatus[0] = cra.getLastRow();
                isInMergedStatus[1] = cra.getLastColumn();
                return isInMergedStatus;
            }
            if (row >= cra.getFirstRow() && row <= cra.getLastRow()) {
                if (col >= cra.getFirstColumn() && col <= cra.getLastColumn()) {
                    isInMergedStatus[0] = 0;
                    isInMergedStatus[1] = 0;
                    return isInMergedStatus;
                }
            }
        }
        return isInMergedStatus;
    }
}
