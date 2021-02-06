package cn.ryoii;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * @author ryoii
 */
public class ExcelCameraConfiguration {

    private String fileName;
    private File file;
    private InputStream inputStream;

    private String sheetName = "Sheet1";

    private double rowZoom = 1.5;
    private double colZoom = 1.5;

    private int rowFrom = 0;
    private int colFrom = 0;
    private int rowTo = 10;
    private int colTo = 10;

    public ExcelCameraConfiguration(String filename) {
        this.fileName = filename;
    }

    public ExcelCameraConfiguration(File file) {
        this.file = file;
    }

    public ExcelCameraConfiguration(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public ExcelCameraConfiguration sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelCameraConfiguration rowZoom(double rowZoom) {
        this.rowZoom = rowZoom;
        return this;
    }

    public ExcelCameraConfiguration colZoom(double colZoom) {
        this.colZoom = colZoom;
        return this;
    }

    public ExcelCameraConfiguration rowFrom(int rowFrom) {
        this.rowFrom = rowFrom;
        return this;
    }

    public ExcelCameraConfiguration colFrom(int colFrom) {
        this.colFrom = colFrom;
        return this;
    }

    public ExcelCameraConfiguration rowTo(int rowTo) {
        this.rowTo = rowTo;
        return this;
    }

    public ExcelCameraConfiguration colTo(int colTo) {
        this.colTo = colTo;
        return this;
    }


    /*
    * internal
    * */

    InputStream createInputStream() throws Exception {
        if (inputStream != null) {
            return inputStream;
        }
        if (file != null) {
            return new FileInputStream(file);
        }
        if (fileName != null) {
            return new FileInputStream(fileName);
        }
        // assert can't reach
        return null;
    }

    String fileName() {
        return fileName;
    }

    File file() {
        return file;
    }

    InputStream inputStream() {
        return inputStream;
    }

    String sheetName() {
        return sheetName;
    }

    double rowZoom() {
        return rowZoom;
    }

    double colZoom() {
        return colZoom;
    }

    int rowFrom() {
        return rowFrom;
    }

    int colFrom() {
        return colFrom;
    }

    int rowTo() {
        return rowTo;
    }

    int colTo() {
        return colTo;
    }
}
