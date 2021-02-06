package cn.ryoii;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import sun.awt.SunHints;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.channels.Channels;
import java.nio.channels.ReadableByteChannel;
import java.util.ArrayList;
import java.util.List;

/**
 * @author ryoii
 */
public class ExcelCamera {

    private final ExcelCameraConfiguration config;
    private List<CellRangeAddress> rangeAddress;
    private final List<Grid> grids;

    private final int rowSize;
    private final int colSize;
    private int imageWidth = 0;
    private int imageHeight = 0;

    private final UserCell[][] cells;

    public ExcelCamera(ExcelCameraConfiguration config) {
        this.config = config;
        this.rowSize = config.rowTo() - config.rowFrom();
        this.colSize = config.colTo() - config.colFrom();
        this.grids = new ArrayList<>(this.colSize * this.rowSize);

        this.cells = new UserCell[rowSize][colSize];
    }

    public File asImageFile(String path) throws Exception {
        File file = new File(path);
        try (InputStream is = asInputStream();
             FileOutputStream fos = new FileOutputStream(file);
             ReadableByteChannel channel = Channels.newChannel(is)) {

            fos.getChannel().transferFrom(channel, 0, is.available());
        }
        return file;
    }

    public InputStream asInputStream() throws Exception {
        ByteArrayOutputStream os = (ByteArrayOutputStream) asOutputStream();
        return new ByteArrayInputStream(os.toByteArray());
    }

    public OutputStream asOutputStream() throws Exception {
        Workbook wb = WorkbookFactory.create(config.createInputStream());
        Sheet sheet = wb.getSheet(config.sheetName());
        if (sheet == null) {
            throw new RuntimeException("Couldn't find sheet " + config.sheetName());
        }
        wb.close();

        // merged regions
        rangeAddress = sheet.getMergedRegions();

        for (int i = 0; i < rowSize; i++) {
            float currentRowWidth = 0;

            // skip empty row
            Row row = sheet.getRow(i + config.rowFrom());
            if (row == null) {
                imageHeight += sheet.getDefaultRowHeightInPoints() * config.colZoom();
                continue;
            }

            // calculate height
            double heightPx = row.getHeightInPoints() * config.colZoom();

            for (int j = 0; j < colSize; j++) {
                UserCell cell = UserCell.build(sheet, i + config.rowFrom(), j + config.colFrom());
                cells[i][j] = cell;

                // calculate width
                double widthPx = sheet.getColumnWidthInPixels(j + config.colFrom()) * config.rowZoom();

                // set cell position
                cell.setTop(imageHeight);
                cell.setBottom((int) (heightPx + imageHeight));
                cell.setLeft((int) currentRowWidth);
                cell.setRight((int) (widthPx + currentRowWidth));

                currentRowWidth += widthPx * config.rowZoom();
            }

            imageHeight += heightPx;
            imageWidth = Math.max(imageWidth, (int) currentRowWidth);
        }

        // gen grid
        for (int i = 0; i < rowSize; i++) {
            for (int j = 0; j < colSize; j++) {
                if (cells[i][j] == null) {
                    continue;
                }
                Grid grid = Grid.build(wb, cells[i][j]);
                grid.setFontZoom(config.colZoom());
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
                    // merge border
                    CellStyle lastCellStyle = cells[lastRowPos][lastColPos].getCell().getCellStyle();
                    if (lastCellStyle.getBorderRight() != BorderStyle.NONE) {
                        grid.setBorderRight();
                    }
                    if (lastCellStyle.getBorderBottom() != BorderStyle.NONE) {
                        grid.setBorderBottom();
                    }
                } // else: empty cell

                grids.add(grid);
            }
        }

        return drawPic();
    }

    private OutputStream drawPic() throws IOException {
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
            if (g.getBgColor() != null) {
                g2d.setColor(g.getBgColor() == null ? java.awt.Color.white : g.getAwtBgColor());
                g2d.fillRect(g.getX(), g.getY(), g.getWidth(), g.getHeight());

            }
            // draw the border
            g2d.setColor(Color.black);
            g2d.setStroke(new BasicStroke(1));

            if (g.hashBorderTop()) {
                g2d.drawLine(g.getX(), g.getY() - 1, g.getX() + g.getWidth() - 1, g.getY() - 1);
            }
            if (g.hashBorderRight()) {
                g2d.drawLine(g.getX() + g.getWidth() - 1, g.getY(), g.getX() + g.getWidth() - 1, g.getY() + g.getHeight() - 1);
            }
            if (g.hashBorderBottom()) {
                g2d.drawLine(g.getX(), g.getY() + g.getHeight() - 1, g.getX() + g.getWidth() - 1, g.getY() + g.getHeight() - 1);
            }
            if (g.hashBorderLeft()) {
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
                switch (g.getAlign()) {
                    case 1:
                        x = g.getX() + (g.getWidth() - strWidth - 1);
                        break;
                    case 0:
                        x = g.getX() + (g.getWidth() - strWidth) / 2 - 1;
                        break;
                    default:
                        x = g.getX();
                        break;
                }
                g2d.drawString(g.getText(), x,
                        g.getY() + (g.getHeight() - font.getSize()) / 2 + font.getSize());
            }
        }

        g2d.dispose();

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", bos);
        return bos;
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
