/*
 * Copyright (c) 2018, Temple University
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * * Redistributions of source code must retain the above copyright notice, this
 *   list of conditions and the following disclaimer.
 * * Redistributions in binary form must reproduce the above copyright notice,
 *   this list of conditions and the following disclaimer in the documentation
 *   and/or other materials provided with the distribution.
 * * All advertising materials features or use of this software must display 
 *   the following  acknowledgement
 *   This product includes software developed by Temple University
 * * Neither the name of the copyright holder nor the names of its 
 *   contributors may be used to endorse or promote products derived 
 *   from this software without specific prior written permission. 
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */
package edu.temple.cla.papolicy.wolfgang.loaddbfromxlsx;

import java.io.InputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import javax.sql.DataSource;
import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import static org.apache.poi.ss.usermodel.CellType.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Paul Wolfgang
 */
public class DoUpload {

    private static final Logger LOGGER = Logger.getLogger(DoUpload.class);
    private final DataSource dataSource;

    public DoUpload(DataSource dataSource) {
        this.dataSource = dataSource;
    }

    public void run(InputStream input) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook(input);
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            Row firstRow = rowIterator.next();
            Map<Integer, String> columnLablesByIndex = new TreeMap<>();
            Map<String, Integer> columnLables = new TreeMap<>();
            firstRow.forEach(cell -> {
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case _NONE:
                    case BLANK:
                    case BOOLEAN:
                    case ERROR:
                    case FORMULA:
                    case NUMERIC:
                    case STRING:
                        int columnIndex = cell.getColumnIndex();
                        String columnValue = cell.getStringCellValue();
                        columnLablesByIndex.put(columnIndex, columnValue);
                        columnLables.put(columnValue, columnIndex);
                        break;
                }
            });
            System.out.println(columnLablesByIndex);
        } catch (Exception e) {
            LOGGER.error("Error processing ", e);
        }
    }

}
