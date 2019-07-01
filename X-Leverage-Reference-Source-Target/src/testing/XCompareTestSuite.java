package testing;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;

@RunWith(Suite.class)

@Suite.SuiteClasses({
	XCompareUnitTest.class,
	XCompareIntegrationTest.class,
})

public class XCompareTestSuite {

}
