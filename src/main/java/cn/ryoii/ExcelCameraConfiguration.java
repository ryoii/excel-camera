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
