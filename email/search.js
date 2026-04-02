/**
 * Search emails functionality
 *
 * Microsoft Graph API constraints:
 * - $search: plain text keywords only, no KQL field:value, no $orderby/$filter
 * - $filter: structured filters on properties like from, hasAttachments, isRead
 * - $filter on from/to requires ConsistencyLevel: eventual header + $count=true
 * - $search and $filter CANNOT be combined
 *
 * Strategy:
 * 1. from/to specified → use $filter with advanced query headers, client-side text filter
 * 2. Only text query/subject → use $search
 * 3. Only boolean filters → use $filter
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

async function handleSearchEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = await resolveFolderPath(accessToken, folder);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder}`);

    const hasStructuredFilter = !!(from || to);
    const hasTextSearch = !!(query || subject);
    const hasBooleanFilter = hasAttachments === true || unreadOnly === true;

    const params = {
      $select: config.EMAIL_SELECT_FIELDS
    };
    let extraHeaders = {};
    let clientSideTextFilter = null;

    if (hasStructuredFilter) {
      // Use $filter with advanced query for from/to
      const filterConditions = [];
      if (from) filterConditions.push(`from/emailAddress/address eq '${from}'`);
      if (to) filterConditions.push(`toRecipients/any(r:r/emailAddress/address eq '${to}')`);
      if (hasAttachments === true) filterConditions.push('hasAttachments eq true');
      if (unreadOnly === true) filterConditions.push('isRead eq false');

      params.$filter = filterConditions.join(' and ');
      params.$count = 'true';
      params.$top = 50;
      extraHeaders = { 'ConsistencyLevel': 'eventual' };

      // Text search will be done client-side
      if (hasTextSearch) {
        clientSideTextFilter = { query, subject };
      }
    } else if (hasTextSearch) {
      // Use $search for text queries (no $orderby allowed)
      const searchTerms = [];
      if (query) searchTerms.push(`"${query}"`);
      if (subject) searchTerms.push(`"${subject}"`);
      params.$search = searchTerms.join(' ');
      params.$top = Math.min(50, requestedCount);

      // Boolean filters done client-side since $search + $filter not allowed
      if (hasBooleanFilter) {
        params.$top = 50;
      }
    } else if (hasBooleanFilter) {
      // Only boolean filters
      const filterConditions = [];
      if (hasAttachments === true) filterConditions.push('hasAttachments eq true');
      if (unreadOnly === true) filterConditions.push('isRead eq false');
      params.$filter = filterConditions.join(' and ');
      params.$orderby = 'receivedDateTime desc';
      params.$top = Math.min(50, requestedCount);
    } else {
      // No criteria — recent emails
      params.$orderby = 'receivedDateTime desc';
      params.$top = Math.min(50, requestedCount);
    }

    console.error("Search params:", JSON.stringify(params, null, 2));
    if (clientSideTextFilter) console.error("Client-side text filter:", JSON.stringify(clientSideTextFilter));
    if (Object.keys(extraHeaders).length > 0) console.error("Extra headers:", JSON.stringify(extraHeaders));

    const fetchCount = (hasStructuredFilter || (hasTextSearch && hasBooleanFilter))
      ? Math.max(requestedCount * 5, 50)
      : requestedCount;

    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, fetchCount, extraHeaders);
    console.error(`API returned ${response.value?.length || 0} results`);

    // Client-side text filtering (when $filter was used instead of $search)
    if (clientSideTextFilter && response.value) {
      const q = (clientSideTextFilter.query || '').toLowerCase();
      const s = (clientSideTextFilter.subject || '').toLowerCase();
      response.value = response.value.filter(email => {
        const emailSubject = (email.subject || '').toLowerCase();
        const emailPreview = (email.bodyPreview || '').toLowerCase();
        if (q && !emailSubject.includes(q) && !emailPreview.includes(q)) return false;
        if (s && !emailSubject.includes(s)) return false;
        return true;
      });
      console.error(`After text filtering: ${response.value.length} results`);
    }

    // Client-side boolean filtering (when $search was used instead of $filter)
    if (hasTextSearch && !hasStructuredFilter && hasBooleanFilter && response.value) {
      response.value = response.value.filter(email => {
        if (hasAttachments === true && !email.hasAttachments) return false;
        if (unreadOnly === true && email.isRead) return false;
        return true;
      });
      console.error(`After boolean filtering: ${response.value.length} results`);
    }

    // Trim to requested count
    if (response.value && response.value.length > requestedCount) {
      response.value = response.value.slice(0, requestedCount);
    }

    return formatSearchResults(response);
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    return {
      content: [{
        type: "text",
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

function formatSearchResults(response) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{
        type: "text",
        text: `No emails found matching your search criteria.`
      }]
    };
  }

  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails matching your search criteria:\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;
