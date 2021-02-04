package cn.ryoii;

/**
 * @author ryoii
 */
public class SnipShotConfig {

    private final double rowZoom;
    private final double colZoom;
    private final String fileName;
    private final String sheetName;
    private final int rowFrom;
    private final int colFrom;
    private final int rowTo;
    private final int colTo;

    public SnipShotConfig(double rowZoom, double colZoom, String fileName, String sheetName, int rowFrom, int colFrom, int rowTo, int colTo) {
        this.rowZoom = rowZoom;
        this.colZoom = colZoom;
        this.fileName = fileName;
        this.sheetName = sheetName;
        this.rowFrom = rowFrom;
        this.colFrom = colFrom;
        this.rowTo = rowTo;
        this.colTo = colTo;
    }

    public static class SnipShotConfigBuilder {

        private double rowZoom = 1.8;
        private double colZoom = 1.8;
        private final String fileName;
        private String sheetName = "Sheet1";
        private int rowFrom = 0;
        private int colFrom = 0;
        private int rowTo = 10;
        private int colTo = 10;

        public SnipShotConfigBuilder(String fileName){
            this.fileName = fileName;
        }

        public SnipShotConfigBuilder rowZoom(double rowZoom) {
            this.rowZoom = rowZoom;
            return this;
        }

        public SnipShotConfigBuilder colZoom(double colZoom) {
            this.colZoom = colZoom;
            return this;
        }

        public SnipShotConfigBuilder sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        public SnipShotConfigBuilder rowFrom(int rowFrom) {
            this.rowFrom = rowFrom;
            return this;
        }

        public SnipShotConfigBuilder colFrom(int colFrom) {
            this.colFrom = colFrom;
            return this;
        }

        public SnipShotConfigBuilder rowTo(int rowTo) {
            this.rowTo = rowTo;
            return this;
        }

        public SnipShotConfigBuilder colTo(int colTo) {
            this.colTo = colTo;
            return this;
        }

        public SnipShotConfig build() {
            return new SnipShotConfig(rowZoom, colZoom, fileName, sheetName, rowFrom, colFrom, rowTo, colTo);
        }
    }

    public double getRowZoom() {
        return rowZoom;
    }

    public double getColZoom() {
        return colZoom;
    }

    public String getFileName() {
        return fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getRowFrom() {
        return rowFrom;
    }

    public int getColFrom() {
        return colFrom;
    }

    public int getRowTo() {
        return rowTo;
    }

    public int getColTo() {
        return colTo;
    }
}
