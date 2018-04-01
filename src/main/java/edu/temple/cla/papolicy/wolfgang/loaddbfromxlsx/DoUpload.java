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

import edu.temple.cla.policydb.dbutilities.ColumnMetaData;
import edu.temple.cla.policydb.dbutilities.DBUtil;
import static edu.temple.cla.policydb.dbutilities.DBUtil.doubleQuotes;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.Clock;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import static java.time.format.DateTimeFormatter.ISO_LOCAL_DATE;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.StringJoiner;
import java.util.TreeMap;
import javax.sql.DataSource;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
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
    private static final long MS_PER_DAY = 24 * 3600 * 1000;
    private static final long BASE_TIME = LocalDate.of(1899, 12, 31).toEpochDay();
    private static final LocalDate FEB28_1900 = LocalDate.of(1900, 2, 28);
    private static final long FEB28_1900_NUM_DAYS = 
            (FEB28_1900.toEpochDay() - BASE_TIME);

    private final DataSource dataSource;
    private Map<String, String> legalToOriginal;
    private List<String> columnNames;

    public DoUpload(DataSource dataSource) {
        this.dataSource = dataSource;
    }

    public void run(InputStream input, String tableName) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook(input);
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            Row firstRow = rowIterator.next();
            Map<String, Integer> columnNamesToIndex = new TreeMap<>();
            columnNames = new ArrayList<>();
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
                        columnNames.set(columnIndex, columnValue);
                        columnNamesToIndex.put(columnValue, columnIndex);
                        break;
                }
            });
            Connection conn = dataSource.getConnection();
            DatabaseMetaData metaData = conn.getMetaData();
            List<ColumnMetaData> columnList;
            try (ResultSet rs2 = metaData.getColumns(null, null, tableName, null)) {
                columnList = ColumnMetaData.getColumnMetaDataList(rs2);
            }
            List<ColumnMetaData> newColumnList = createNewColumnList(columnList);
            
        } catch (Exception e) {
            LOGGER.error("Error processing ", e);
        }
    }

    List<ColumnMetaData> createNewColumnList(List<ColumnMetaData> columnList) {
        legalToOriginal = new HashMap<>();
        List<ColumnMetaData> newColumnList = new ArrayList<>();
        columnNames.forEach((columnName) -> {
            String dbColumnName = DBUtil.convertToLegalName(columnName).toString();
            legalToOriginal.put(dbColumnName, columnName);
        });
        columnList.stream()
                .filter((metaData)-> (legalToOriginal.get(metaData.getColumnName()) != null))
                .forEach(newColumnList::add);
        return newColumnList;
    }

        public Optional<String> buildValuesList(Row row,
            List<ColumnMetaData> metaDataList,
            List<ColumnMetaData> newColumnList) throws SQLException, Exception {
        Map<String, String> record = new HashMap<>();
        for (Cell cell : row) {
            int columnIndex = cell.getColumnIndex();
            String value = null;
            CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case _NONE: break;
                    case BLANK: break;
                    case BOOLEAN: 
                        value = Boolean.toString(cell.getBooleanCellValue());
                        break;
                    case ERROR: break;
                    case FORMULA: break;
                    case NUMERIC:
                        value = Double.toString(cell.getNumericCellValue());
                        break;
                    case STRING:
                        value = cell.getStringCellValue();
                        break;
                    default:
                        break;
                }
            if (value != null) {
                record.put(columnNames.get(columnIndex), value);
            }
            columnIndex++;
        }
        if (record.isEmpty()) return Optional.empty();
        StringJoiner valuesList = new StringJoiner(", ", "(", ")");
        for (ColumnMetaData metaData : metaDataList) {
            int columnType = metaData.getDataType();
            String columnName = metaData.getColumnName();
            String originalColumnName = legalToOriginal.get(columnName);
            if (originalColumnName != null) {
                String value = record.get(originalColumnName);
                if (value == null || value.isEmpty()) {
                    valuesList.add("NULL");
                } else {
                    switch (columnType) {
                        case java.sql.Types.BINARY:
                        case java.sql.Types.VARBINARY:
                            valuesList.add(value);
                            break;
                        case java.sql.Types.CHAR:
                        case java.sql.Types.VARCHAR:
                        case java.sql.Types.LONGVARCHAR:
                            valuesList.add("\'" + doubleQuotes(value) + "\'");
                            break;
                        case java.sql.Types.REAL:
                        case java.sql.Types.DOUBLE:
                            valuesList.add(DBUtil.removeCommas(value));
                            break;
                        case java.sql.Types.BIT:
                        case java.sql.Types.SMALLINT:
                        case java.sql.Types.INTEGER:
                            valuesList.add(removeFraction(value));
                            break;
                        case java.sql.Types.TIMESTAMP:
                        case java.sql.Types.DATE:
                            valuesList.add("'" + value + "'");
                            break;
                        default:
                            System.err.println("Unrecognized type: " + columnType);
                            System.err.println("Type name: " + columnName);
                            throw new Exception();
                    }
                }
            }
        }
        return Optional.of(valuesList.toString());
    }
        
    public static String removeFraction(String number) {
        int posDot = number.indexOf(".");
        return number.substring(posDot);
    }
    
    public static String excelDateToDate(String excelDateString) {
        double excelDate = Double.parseDouble(excelDateString);
        long numberOfDays = (long)excelDate;
        double partOfDay = excelDate - numberOfDays;
        if (numberOfDays > FEB28_1900_NUM_DAYS) {
            numberOfDays--;
        }
        long javaDateValue = (numberOfDays + BASE_TIME)*MS_PER_DAY + (long)(partOfDay * MS_PER_DAY);
        Instant javaInstant = Instant.ofEpochMilli(javaDateValue);
        LocalDateTime javaDateTime = LocalDateTime.ofInstant(javaInstant, ZoneOffset.UTC);
        String javaDateString = javaDateTime.format(ISO_LOCAL_DATE);
        return javaDateString;
    }
    
}
