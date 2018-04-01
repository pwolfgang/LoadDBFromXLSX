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

/**
 *
 * @author Paul Wolfgang
 */
public class Main {
    
    public static void main(String[] args) throws Exception {
        DataSource dataSource = new SimpleDataSource(args[0]);
        DoUpload doUpload = new DoUpload(dataSource);
        InputStream input = new FileInputStream(args[2]);
        doUpload.run(input, args[1]);
    }    
}
