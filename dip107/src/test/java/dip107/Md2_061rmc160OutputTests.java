package dip107;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvFileSource;
import org.junit.jupiter.params.provider.ValueSource;

public class Md2_061rmc160OutputTests {
    private ByteArrayOutputStream byteArrayOutputStream;
    private String ObjectUnderTestName = "dip107.Md2_061rmc160";

    @ParameterizedTest
    @ValueSource(floats = { 1 })
    public void shouldPrintAplnrVardsUzvardsGrupasNr(float input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(input + ""), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        assertEquals("061RMC160 Oskars Grauzis 4", output[0]);
    }

    @ParameterizedTest
    @ValueSource(floats = { 1, 2, -3, -4 })
    public void shouldPrintNumberPrompt(float input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(input + ""), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        Boolean hasOutItem = output.length > 1;
        if (hasOutItem)
            assertEquals("K=", output[1]);
        else
            assertTrue(false,
                    "the parameter prompt should be the second line outputted! Output had lines: " + output.length);
    }

    @ParameterizedTest
    @ValueSource(floats = { 1 })
    public void shouldPrintResultTitle(float input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        String t = Float.toString(input);
        runTest(getSimulatedUserInput(t), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        Boolean hasOutItem = output.length > 2;
        if (hasOutItem)
            // TODO: kaa jaabuut? "a=result:"-> ja in dos \r\n "result:"-> ja formateshu lai
            // butu \r\n...
            assertEquals("result:", output[2]);
        else
            assertTrue(false,
                    "the result: text should be the third line outputted! Output had lines: " + output.length
                            + System.getProperty("line.separator") + "output: " + byteArrayOutputStream.toString()
                            + "input: " + t);
    }

    @ParameterizedTest
    @ValueSource(floats = { 1 })
    public void shouldPrintArrayNames(float input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(input + ""), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        // TODO: kaa jaabuut? 2-> ja in dos \r\n 3-> ja formateshu lai butu \r\n...
        Boolean hadSecondName = false;
        Boolean hasOutItem = output.length > 3;
        if (hasOutItem)
            assertTrue(output[3].matches("[ABC]:"), "should have array name followed by column!");
        else
            assertTrue(false,
                    "the array title should be the fourth line outputted! Output had lines: " + output.length);

        for (int i = 3; i < output.length; i++) {
            if (output[i].matches("[ABC]:")) {
                hadSecondName = true;
                break;
            }
        }
        assertTrue(hadSecondName, "Should output two array titles!");
    }

    @ParameterizedTest
    @ValueSource(floats = { 1 })
    public void shouldPrintFormattedResults(float input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(input + ""), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        Boolean hasOutItems = output.length > 12;
        if (hasOutItems)
            for (int i = 4; i <= 12; i++) {
                // 2. masīva virsraksts
                if (i == 7)
                    continue;
                assertTrue(output[i].matches("([+-]?\\d+[\\.,]\\d{2}\\s+){4}([+-]?\\d+[\\.,]\\d{2})"),
                        "Line number " + (i + 1) + System.getProperty("line.separator")
                                + "Should output result: 5 numbers per line with 2 decimal places!"
                                + System.getProperty("line.separator") + "output was: " + output[i]
                                + System.getProperty("line.separator"));
            }
        else
            assertTrue(false,
                    "the results output should start at the fifith line outputted and have two sets of 4 rows of properly formatted results! Output had lines: "
                            + output.length + System.getProperty("line.separator") + "output was: "
                            + byteArrayOutputStream.toString() + System.getProperty("line.separator"));
    }

    @ParameterizedTest
    @ValueSource(floats = { 1 })
    public void shouldPrintFormattedResultsAndSortedFirstArray(float input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(input + ""), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        Boolean hasOutItems = output.length > 12;
        String pattern = "([+-]?\\d+[\\.,]\\d{2}\\s+)([+-]?\\d+[\\.,]\\d{2}\\s+)([+-]?\\d+[\\.,]\\d{2}\\s+)([+-]?\\d+[\\.,]\\d{2}\\s+)([+-]?\\d+[\\.,]\\d{2})";
        if (hasOutItems)
            for (int i = 4; i <= 12; i++){
                //2. masīva virsraksts
                if(i==8) continue;
                //assertTrue(output[i].matches("([+-]?\\d+[\\.,]\\d{2}\\s+){4}([+-]?\\d+[\\.,]\\d{2})"),
                assertTrue(output[i].matches(pattern),
                        "Line number " + (i + 1) + System.getProperty("line.separator")
                                + "Should output result: 5 numbers per line with 2 decimal places!"
                                + System.getProperty("line.separator") + "output was: " + output[i]
                                + System.getProperty("line.separator"));
            } else
            assertTrue(false,
                    "the results output should start at the fifith line outputted and have two sets of 4 rows of properly formatted results! Output had lines: "
                            + output.length
                            +System.getProperty("line.separator")
                            +"output was: "+byteArrayOutputStream.toString()
                            +System.getProperty("line.separator"));

                            //region parse
                            float aArray[][] = new float[4][5];
                            float bArray[][] = new float[4][5];
                            Pattern r = Pattern.compile(pattern);
                            Matcher m;
                            for (int i=4; i<8; i++){
                                m=r.matcher(output[i]);
                                //paarbaude jau bija (ka deriigi floati!) bet tāpat jāizsauc, lai grupas saformē - ielasa..
                                m.find();
                                for(int j=0; j<m.groupCount(); j++)
                                    aArray[i-4][j]= Float.parseFloat(m.group(j));
                            }   
                            for (int i=9; i<=12; i++){
                                m=r.matcher(output[i]);
                                //paarbaude jau bija (ka deriigi floati!) bet tāpat jāizsauc, lai grupas saformē - ielasa..
                                m.find();
                                for(int j=0; j<m.groupCount(); j++)
                                    bArray[i-9][j]= Float.parseFloat(m.group(j));
                            }         
                            //endregion
                            
                             
                            //region prepareCheck
                            List<Float> listPos=new ArrayList<Float>(); 
                            List<Float> listNeg=new ArrayList<Float>();
                            for(int i=0; i<4; i++){
                                for(int j=0;j<5;j++){
                                    //kas ar nulli?? Pozitīva?
                                    if(aArray[i][j]>=0) listPos.add(aArray[i][j]);
                                    else listNeg.add(aArray[i][j]);
                                }
                            } 
                            //endregion

                            //region check
                            int k=listPos.size(), cnt=0;
                            for(int i=0; i<4; i++){
                                for(int j=0;j<5;j++){
                                    assertEquals(cnt<k? listPos.get(cnt) : listNeg.get(cnt), bArray[i][j], 
                                    "The list B should be sorted list A according to specification!"
                                    +System.getProperty("line.separator")
                                    +(cnt >1 ? "last two elements: "+(cnt-2<k? listPos.get(cnt-2) : listNeg.get(cnt-2))
                                    + " and "+(cnt-1<k? listPos.get(cnt-1) : listNeg.get(cnt-1)) 
                                    +System.getProperty("line.separator")
                                    +" should be followed by "+ (cnt<k? listPos.get(cnt) : listNeg.get(cnt))
                                    +System.getProperty("line.separator")
                                    + " but was followed by " + bArray[i][j]
                                    :""));
                                    cnt++;
                                }
                            } 
                            //endregion
    }

    @ParameterizedTest
    @ValueSource(strings = { "fas", "-+1", "š", "0.0.0.0", "8k8" })
    public void shouldTellWrongInput(String input) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(input + ""), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        if (output.length > 1)
            assertEquals("input-output error", output[output.length - 1],
                    "on error should output 'input-output error'");
        else
            assertTrue(false, "the program should output at least one line! Output had lines: " + output.length);
    }

    // endregion
    // region utils
    private String getSimulatedUserInput(String... inputs) {
        return String.join(System.getProperty("line.separator"), inputs) + System.getProperty("line.separator");
    }

    private void runTest(String data, String className) throws Exception {

        InputStream input = new ByteArrayInputStream(data.getBytes("UTF-8"));
        ;

        Class<?> cls = Class.forName(className);
        Object t = cls.getDeclaredConstructor().newInstance();
        Method meth = t.getClass().getDeclaredMethod("testableMain", InputStream.class, PrintStream.class);

        meth.invoke(t, input, new PrintStream(byteArrayOutputStream));
    }
    // endregion
}
