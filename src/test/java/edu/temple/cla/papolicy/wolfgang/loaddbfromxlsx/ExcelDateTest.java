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

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.TimeZone;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

/**
 *
 * @author Paul
 */
@RunWith(Parameterized.class)
public class ExcelDateTest {
    
    @Parameters
    public static Collection<Object[]> data() {
        List<Object[]> params = 
                new ArrayList<>(Arrays.asList(new Object[][]{
                    {"1900-01-01", 1},
                    {"1900-01-02", 2},
                    {"1900-01-31", 31},
                    {"1900-02-01", 32},
                    {"1900-02-28", 59},
                    {"1900-03-01", 61},
                    {"1900-07-10", 192},
                    {"1901-01-01", 367},
                    {"1901-01-02", 368},
                    {"1901-01-31", 397},
                    {"1901-02-01", 398},
                    {"1901-02-28", 425},
                    {"1901-03-01", 426},
                    {"1901-07-10", 557},
                    {"1979-01-01", 28856},
                    {"1979-01-02", 28857},
                    {"1979-01-31", 28886},
                    {"1979-02-01", 28887},
                    {"1979-02-28", 28914},
                    {"1979-03-01", 28915},
                    {"1979-03-02", 28916},
                    {"1979-07-09", 29045},
                    {"1979-07-10", 29046}            
        }));
        GregorianCalendar calendar1979_03_01 = new GregorianCalendar();
        calendar1979_03_01.setTimeZone(TimeZone.getTimeZone("GMT"));
        calendar1979_03_01.clear();
        calendar1979_03_01.set(1979, 2, 1);
        Date date1979_03_01 = calendar1979_03_01.getTime();
        long goodDateFor1979_03_01 = 28915;
        long msPerDay = 24 * 3600 * 1000;
        for (int i = 0; i < 150; i++) {
            Date d = new Date(date1979_03_01.getTime() + i * msPerDay);
            String dateString = String.format("%tF", d.getTime()+ 5*3600*1000);
            Long expected = goodDateFor1979_03_01 + i;
            params.add(new Object[]{dateString, expected});
        }
        return params;
    }
    
    private final double input;
    private final String expected;
    
    public ExcelDateTest(String expected, double input) {
        this.input = input;
        this.expected = expected;
    }

    @Test
    public void testExcelDateToDate() {
        System.out.println("excelDateToDate");
        String excelDateString = Double.toString(input);
        String expResult = expected;
        String result = DoUpload.excelDateToDate(excelDateString);
        assertEquals(expResult, result);
    }
    
}
