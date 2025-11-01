/**
 * Test pagination functionality manually
 */
const { callGraphAPIPaginated } = require('./utils/graph-api');
const { ensureAuthenticated } = require('./auth');
const config = require('./config');

/**
 * Helper function to run a single pagination test
 * @param {string} accessToken - Access token for authentication
 * @param {string} description - Description of the test
 * @param {number} maxCount - Maximum number of emails to retrieve
 * @param {number} pageSize - Page size for API calls
 * @returns {Promise<{result: object, duration: number}>} - Test results
 */
async function runTest(accessToken, description, maxCount, pageSize) {
  console.log(description);
  const start = Date.now();
  const testResult = await callGraphAPIPaginated(
    accessToken,
    'GET',
    'me/messages',
    {
      $top: pageSize,
      $orderby: 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    },
    maxCount
  );
  const duration = Date.now() - start;
  console.log(`   ✅ Retrieved ${testResult.value.length} emails in ${duration}ms\n`);
  return { result: testResult, duration };
}

async function testPagination() {
  try {
    console.log('🧪 Testing Pagination Functionality\n');
    console.log('===================================\n');

    // Get access token
    console.log('1. Authenticating...');
    const accessToken = await ensureAuthenticated();
    console.log('   ✅ Authentication successful\n');

    // Run tests with varying counts
    const { result: test1 } = await runTest(
      accessToken,
      '2. Test 1: Fetching 10 emails (single page)...',
      10,
      10
    );

    const { result: test2, duration: test2Duration } = await runTest(
      accessToken,
      '3. Test 2: Fetching 100 emails (requires pagination)...',
      100,
      config.API_PAGE_SIZE
    );

    const { result: test3, duration: test3Duration } = await runTest(
      accessToken,
      '4. Test 3: Fetching 200 emails (extensive pagination)...',
      200,
      config.API_PAGE_SIZE
    );

    // Summary
    console.log('===================================');
    console.log('✅ All pagination tests passed!');
    console.log(`\nResults:`);
    console.log(`  - 10 emails: ${test1.value.length} retrieved`);
    console.log(`  - 100 emails: ${test2.value.length} retrieved (${test2Duration}ms)`);
    console.log(`  - 200 emails: ${test3.value.length} retrieved (${test3Duration}ms)`);

    // Show sample of latest emails
    console.log(`\n📧 Sample (Latest 3 emails):`);
    test1.value.slice(0, 3).forEach((email, i) => {
      const date = new Date(email.receivedDateTime).toLocaleDateString();
      const from = email.from?.emailAddress?.name || 'Unknown';
      console.log(`   ${i + 1}. [${date}] ${from}: ${email.subject}`);
    });

  } catch (error) {
    // Log the full error object so non-Error throwables are readable
    console.error('❌ Test failed:', error);
    // Print stack only when available
    if (error && error.stack) {
      console.error(error.stack);
    }
    // Prefer setting exit code to allow graceful cleanup by other hooks
    process.exitCode = 1;
  }
}

// Run tests
testPagination();
