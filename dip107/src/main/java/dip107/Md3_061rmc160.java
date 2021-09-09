package dip107;

import java.io.InputStream;
import java.io.PrintStream;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.Scanner;
import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * Hello world!
 *
 */
public class Md3_061rmc160 {
    // omg.. SERIOUSLY???
    // https://stackoverflow.com/questions/25179127/multiplying-two-numbers-using-bigdecimal-returns-a-wrong-value/25179147#25179147
    // and also a good info:
    // https://stackoverflow.com/questions/5384601/java-doubles-are-not-good-at-math/5385202#5385202
    // what every computer scientist should know about floating-point arithmetic...
    // region utils
    private static String lastRawFormattedInput = "";

    private static String makeFloatString(String inputString) {
        DecimalFormat format = (DecimalFormat) DecimalFormat.getInstance();
        DecimalFormatSymbols symbols = format.getDecimalFormatSymbols();
        char sep = symbols.getDecimalSeparator();
        if (inputString.indexOf(sep) > -1)
            return inputString;
        else {
            char otherSep = sep == ',' ? '.' : ',';
            return inputString.replace(otherSep, sep);
        }
    }

    // no passing by reference possible in Java at all?? aww...
    private static float getInput(Scanner sc, PrintStream outputStream, char varName) {
        outputStream.print(varName + "=");
        // wtf-cross thread write happening?? private to class not object - omg wth
        system_exit = false;
        // infinity is an invalid value legal float value example for coordinates!
        if (sc.hasNext("[+-]?[\\d]+([\\.,]\\d+)?")) {
            lastRawFormattedInput = makeFloatString(sc.next());
            return Float.parseFloat(lastRawFormattedInput);
        } else {
            sc.next();
            outputStream.println();
            outputStream.println("input-output error");
            system_exit = true;
            return -11111111.222222f;
        }
    }
    // endregion

    // region main
    public static void main(String[] args) {
        testableMain(System.in, System.out);
    }

    private static Boolean system_exit = false;

    public static void testableMain(InputStream inputStream, PrintStream outputStream) {
        Scanner sc = new Scanner(inputStream);
        // https://blogs.oracle.com/corejavatechtips/the-need-for-bigdecimal
        // Excel and calculator different values!!!!! Unable to get guaranteed precise
        // test data
        // Therefore unable to verify!!!!
        BigDecimal K;
        BigDecimal A[] = new BigDecimal[20];
        String outputFormatString = "%1$.2f";
        // BigDecimal amount = new BigDecimal("100.05");
        outputStream.println("061RMC160 Oskars Grauzis 4");
        getInput(sc, outputStream, 'K');
        // TODO: pajautaat kaa jaabuut - ka enter no usera (un steramaa taatad kopa
        // prompt a=result:)
        outputStream.println();
        if (system_exit) {
            sc.close();
            return;
        }
        K = new BigDecimal(lastRawFormattedInput);
        // region 1. inicializācija
        if (K.compareTo(new BigDecimal("0")) >= 0) {
            Random r = new Random();
            int i = 0;
            while (i < 20) {
                A[i] = new BigDecimal(r.nextDouble() * 100 - 50);
                i++;
            }
        } else {
            A[0] = new BigDecimal("0.1");
            int i = 1;
            while (i < 20) {
                A[i] = A[i - 1].multiply(new BigDecimal(lastRawFormattedInput));
                i++;
            }
        }
        // endregion
        // region 2. izvade
        outputStream.println("result:");
        outputStream.println("A:");
        int i = 0;
        do {
            if ((i % 5) == 4) {
                outputStream.print(String.format(outputFormatString, A[i]) + System.getProperty("line.separator"));
            } else {
                outputStream.print(String.format(outputFormatString, A[i]) + "\t");
            }
            i++;
        } while (i < 20);
        // endregion
        // region 3. apstrāde
        BigDecimal B[] = new BigDecimal[20];
        List<BigDecimal> posList = new ArrayList<BigDecimal>();
        List<BigDecimal> negList = new ArrayList<BigDecimal>();
        for (i = 0; i < 20; i++) {
            if (A[i].compareTo(new BigDecimal("0")) >= 0)
                posList.add(A[i]);
            else
                negList.add(A[i]);
        }
        posList.addAll(negList);
        for (i = 0; i < 20; i++)
            B[i] = posList.get(i);

        // endregion
        // region 4. izvade
        outputStream.println("B:");
        for (i = 0; i < 20; i++) {
            if ((i % 5) == 4) {
                outputStream.print(String.format(outputFormatString, B[i]) + System.getProperty("line.separator"));
            } else {
                outputStream.print(String.format(outputFormatString, B[i]) + "\t");
            }
        }
        // endregion
        // Trešais no beigām studenta apliecības numura cipars 1 vai 6: while/ do while/
        // for /for

    }
    // endregion
}
