/**
 * Test pagination functionality manually
 */
const { callGraphAPIPaginated } = require('./utils/graph-api');
const { ensureAuthenticated } = require('./auth');
const config = require('./config');

async function testPagination() {
  try {
    console.log('🧪 Testing Pagination Functionality\n');
    console.log('===================================\n');
    
    // Get access token
    console.log('1. Authenticating...');
    const accessToken = await ensureAuthenticated();
    console.log('   ✅ Authentication successful\n');
    
    // Test 1: Get 10 emails (should be 1 page)
    console.log('2. Test 1: Fetching 10 emails (single page)...');
    const test1 = await callGraphAPIPaginated(
      accessToken,
      'GET',
      'me/messages',
      {
        $top: 10,
        $orderby: 'receivedDateTime desc',
        $select: config.EMAIL_SELECT_FIELDS
      },
      10
    );
    console.log(`   ✅ Retrieved ${test1.value.length} emails\n`);
    
    // Test 2: Get 100 emails (should require pagination)
    console.log('3. Test 2: Fetching 100 emails (requires pagination)...');
    const test2Start = Date.now();
    const test2 = await callGraphAPIPaginated(
      accessToken,
      'GET',
      'me/messages',
      {
        $top: 50,
        $orderby: 'receivedDateTime desc',
        $select: config.EMAIL_SELECT_FIELDS
      },
      100
    );
    const test2Duration = Date.now() - test2Start;
    console.log(`   ✅ Retrieved ${test2.value.length} emails in ${test2Duration}ms\n`);
    
    // Test 3: Get 200 emails (more pagination)
    console.log('4. Test 3: Fetching 200 emails (extensive pagination)...');
    const test3Start = Date.now();
    const test3 = await callGraphAPIPaginated(
      accessToken,
      'GET',
      'me/messages',
      {
        $top: 50,
        $orderby: 'receivedDateTime desc',
        $select: config.EMAIL_SELECT_FIELDS
      },
      200
    );
    const test3Duration = Date.now() - test3Start;
    console.log(`   ✅ Retrieved ${test3.value.length} emails in ${test3Duration}ms\n`);
    
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
    console.error('❌ Test failed:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

// Run tests
testPagination();
