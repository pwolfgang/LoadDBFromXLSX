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
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
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
import java.util.Objects;
import java.util.Optional;
import java.util.StringJoiner;
import java.util.stream.Collectors;
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
    private static final long FEB28_1900_NUM_DAYS
            = (FEB28_1900.toEpochDay() - BASE_TIME);

    private final DataSource dataSource;
    private Map<String, String> databaseToSpteadsheetNames;
    private List<String> spreadsheetColumnNames;

    /**
     * Constructor.
     *
     * @param dataSource The dataSource referencing the database.
     */
    public DoUpload(DataSource dataSource) {
        this.dataSource = dataSource;
    }

    /**
     * Main program.
     *
     * @param input Input stream containing the xlsx file.
     * @param sheetName Worksheet name containing the data.
     * @param tableName Name of the destination table.
     */
    public void run(InputStream input, String sheetName, String tableName) {
        try (XSSFWorkbook wb = new XSSFWorkbook(input);
                Connection conn = dataSource.getConnection();
                Statement stmt = conn.createStatement();) {
            XSSFSheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                throw new Exception("Sheet " + sheetName + " does not exist");
            }
            Iterator<Row> rowIterator = sheet.iterator();
            Row firstRow = rowIterator.next();
            getSpreadsheetColumnNames(firstRow);
            DatabaseMetaData metaData = conn.getMetaData();
            List<ColumnMetaData> databaseColumnMetadataList;
            try (ResultSet rs2 = metaData.getColumns(null, null, tableName, null)) {
                databaseColumnMetadataList = ColumnMetaData.getColumnMetaDataList(rs2);
            }
            List<ColumnMetaData> filteredColumnList = filterColumnList(databaseColumnMetadataList);
            while (rowIterator.hasNext()) {
                StringJoiner values = new StringJoiner(",\n");
                while (rowIterator.hasNext() && values.length() < 10000000) {
                    Row row = rowIterator.next();
                    buildValuesList(row, filteredColumnList)
                            .ifPresent(values::add);
                }
                String sqlInsertStatement = DBUtil.buildSqlInsertStatement(tableName, filteredColumnList);
                String insert = sqlInsertStatement + "\n" + values.toString();
                try {
                    stmt.executeUpdate(insert);
                } catch (SQLException sqlex) {
                    System.err.println("Error in SQL");
                    System.err.println(insert);
                    throw sqlex;
                }
            }
        } catch (IOException ioex) {
            LOGGER.error("Unable to open workbook", ioex);
        } catch (SQLException sqlex) {
            LOGGER.error("Error accessing database", sqlex);
        } catch (Exception e) {
            LOGGER.error("Error processing ", e);
        }
    }

    public void getSpreadsheetColumnNames(Row firstRow) {
        spreadsheetColumnNames = new ArrayList<>();
        firstRow.forEach(cell -> {
            try {
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
                        spreadsheetColumnNames.add(columnIndex, columnValue);
                        break;
                }
            } catch (Exception ex) {
                throw new RuntimeException("Error processing first row, column "
                        + cell.getColumnIndex(), ex);
            }
        });
    }

    /**
     * Method to create the database column list from the spreadsheet columns
     * and filters our any spreadsheet column that does not have a corresponding
     * database column. This name translation is included since spreadsheets may
     * contain names from an Access database that do not represent legal MySQL
     * names. As a side-effect the Map databaseToSpteadsheetNames is
     * initialized.
     *
     * @param columnList The list of spreadsheet column names.
     * @return The filtered list of legal database column names.
     */
    List<ColumnMetaData> filterColumnList(List<ColumnMetaData> columnList) {
        databaseToSpteadsheetNames = new HashMap<>();
        List<ColumnMetaData> newColumnList = new ArrayList<>();
        spreadsheetColumnNames.forEach((columnName) -> {
            String dbColumnName = DBUtil.convertToLegalName(columnName).toString();
            databaseToSpteadsheetNames.put(dbColumnName, columnName);
        });
        return columnList.stream()
                .filter(metadata -> Objects.nonNull(databaseToSpteadsheetNames.get(metadata.getColumnName())))
                .collect(Collectors.toList());
    }

    /**
     * Method to build the list of values for a row.
     *
     * @param row The spreadsheet row.
     * @param metaDataList The list of column metadata for each database column.
     * @return
     */
    public Optional<String> buildValuesList(Row row,
            List<ColumnMetaData> metaDataList) {
        Map<String, String> record = new HashMap<>();
        for (Cell cell : row) {
            try {
                int columnIndex = cell.getColumnIndex();
                String value = null;
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case _NONE:
                        break;
                    case BLANK:
                        break;
                    case BOOLEAN:
                        value = Boolean.toString(cell.getBooleanCellValue());
                        break;
                    case ERROR:
                        break;
                    case FORMULA:
                        LOGGER.error("Cell " + cell.getAddress() + " contains a formula");
                        break;
                    case NUMERIC:
                        value = String.format("%.0f", cell.getNumericCellValue());
                        break;
                    case STRING:
                        value = cell.getStringCellValue();
                        break;
                    default:
                        break;
                }
                if (value != null) {
                    record.put(spreadsheetColumnNames.get(columnIndex), value);
                }
                columnIndex++;
            } catch (Exception ex) {
                String message = String.format("Error processing row: %d, column: %d",
                        cell.getRowIndex(), cell.getColumnIndex());
                throw new RuntimeException(message, ex);
            }
        }
        if (record.isEmpty()) {
            return Optional.empty();
        }
        StringJoiner valuesList = new StringJoiner(", ", "(", ")");
        for (ColumnMetaData metaData : metaDataList) {
            int columnType = metaData.getDataType();
            String columnName = metaData.getColumnName();
            String spreadsheetColumnName = databaseToSpteadsheetNames.get(columnName);
            if (spreadsheetColumnName != null) {
                String value = record.get(spreadsheetColumnName);
                if (value == null || value.isEmpty() || value.equals("null")) {
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
                        case java.sql.Types.TINYINT:
                        case java.sql.Types.SMALLINT:
                        case java.sql.Types.INTEGER:
                            valuesList.add(removeFraction(value));
                            break;
                        case java.sql.Types.TIMESTAMP:
                        case java.sql.Types.DATE:
                            valuesList.add("'" + excelDateToDate(value) + "'");
                            break;
                        default:
                            String message = "Unrecognized type: " + columnType
                                    + " Type name: " + columnName;
                            return Optional.empty();
                    }
                }
            }
        }
        return Optional.of(valuesList.toString());
    }

    public static String removeFraction(String number) {
        int posDot = number.indexOf(".");
        if (posDot == -1) {
            return number;
        }
        return number.substring(0, posDot);
    }

    public static String excelDateToDate(String excelDateString) {
        double excelDate = Double.parseDouble(excelDateString);
        long numberOfDays = (long) excelDate;
        if (numberOfDays > FEB28_1900_NUM_DAYS) {
            numberOfDays--;
        }
        long javaDateValue = (numberOfDays + BASE_TIME) * MS_PER_DAY;
        Instant javaInstant = Instant.ofEpochMilli(javaDateValue);
        LocalDateTime javaDateTime = LocalDateTime.ofInstant(javaInstant, ZoneOffset.UTC);
        String javaDateString = javaDateTime.format(ISO_LOCAL_DATE);
        return javaDateString;
    }

}
