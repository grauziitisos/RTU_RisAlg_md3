package dip107;

import static org.junit.jupiter.api.Assertions.assertEquals;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.lang.reflect.Method;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvFileSource;

public class Md2_061rmc160ValueTests {

    private ByteArrayOutputStream byteArrayOutputStream;
    private String ObjectUnderTestName = "dip107.Md2_061rmc160";

    // region valTests
    @ParameterizedTest
    @CsvFileSource(resources = "positive-tests.csv", numLinesToSkip = 1)
    void shouldPassAllExcelResultsMappingTests(String x, String expect) throws Exception {
        byteArrayOutputStream = new ByteArrayOutputStream();
        runTest(getSimulatedUserInput(x), ObjectUnderTestName);
        String[] output = byteArrayOutputStream.toString().split(System.getProperty("line.separator"));
        String[] expected = expect.split(System.getProperty("line.separator"));
        // expected ietver arii titles.. (3 un 8 rindas taatad vieglaak testu uzrakstiit
        // :))
        assertEquals(output.length, 13, "The program should output exactly 13 lines to pass this test.. Sorry, no space for creative beautiful designs...");
        assertEquals(expected.length, 10, "Malformed input data! Please check that it is 10 lines total..");
        for (int i = 3; i <= 12; i++)
            assertEquals(expected[i - 3], output[i]);
    }

    // endregion
    // region utils
    private String getSimulatedUserInput(String... inputs) {
        return String.join(System.getProperty("line.separator"), inputs) + System.getProperty("line.separator");
    }

    private void runTest(final String data, final String className) throws Exception {

        final InputStream input = new ByteArrayInputStream(data.getBytes("UTF-8"));
        ;

        final Class<?> cls = Class.forName(className);
        final Method meth = cls.getDeclaredMethod("testableMain", InputStream.class, PrintStream.class);
        meth.invoke(null, input, new PrintStream(byteArrayOutputStream));
    }
    // endregion
}
