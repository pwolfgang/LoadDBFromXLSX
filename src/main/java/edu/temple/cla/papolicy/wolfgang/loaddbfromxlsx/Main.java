/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package edu.temple.cla.papolicy.wolfgang.loaddbfromxlsx;

import edu.temple.cla.policydb.dbutilities.SimpleDataSource;
import java.io.FileInputStream;
import java.io.InputStream;
import javax.sql.DataSource;
import org.apache.log4j.Logger;

/**
 * Standalone program to load a database table from an Excel xlsx file. The
 * class DoUpload can also be used called from the web server application. The
 * first row of the worksheet must contain column labels corresponding to the
 * database columns. The database table must already exist and the data types
 * are determined from the database metadata. The cells must not contain
 * formulae. Cells that contain formula will be treated as NULL. If the sheet
 * contain formulae copy the sheet to a blank sheet using paste-values. This
 * program has only been tested with MySQL database targets, but earlier
 * versions of the code were used to load a Microsoft SQLServer database.
 *
 * @author Paul Wolfgang
 */
public class Main {

    private static final Logger LOGGER = Logger.getLogger(Main.class);

    /**
     * Main class.
     *
     * @param args Command line arguments.
     * <dl>
     * <dt>args[0]</dt><dd>Text file containing DataSource parameters.</dd>
     * <dt>args[1]</dt><dd>The name of the table.</dd>
     * <dt>args[2]</dt><dd>The name of the xlsx file.</dd>
     * <dt>args[3]</dt><dd>The name of the sheet in the workbook containing the
     * data</dd>
     * </dl>
     */
    public static void main(String[] args) {
        try {
            DataSource dataSource = new SimpleDataSource(args[0]);
            DoUpload doUpload = new DoUpload(dataSource);
            InputStream input = new FileInputStream(args[2]);
            doUpload.run(input, args[3], args[1]);
        } catch (Exception ex) {
            LOGGER.error("Error occured", ex);
        }
    }
}
