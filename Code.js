function getStatusText(statusCode) {
  const statusMap = {
    2: "Open",
    3: "Pending",
    4: "Resolved",
    5: "Closed",
    8: "In Progress",
    12: "Clarification - Client/Third Party",
    13: "Tech Dependent",
    14: "Reopened",
    15: "UAT",
    16: "On Hold",
    17: "Temporarily Closed",
    18: "Under Review",
    19: "Rejected",
    20: "New Requirement",
    21: "Confirmation to Closure"
  };
  return statusMap[statusCode] || statusCode;
}

function getAgentName(agentId) {
  const agentMap = {
    17000589198: "Niraj Patil",
    17050112354: "Priyanka Diwate",
    17051657819: "Harshal Bothara",
    17055595704: "Tejas Shivaji Patil",
    17057807143: "Raj Amrutakar",
    17058119097: "Zealkumar Mohodkar",
    17066388956: "Pratik Dahale"
  };
  return agentId ? (agentMap[agentId] || agentId) : "Unassigned";
}

function updateFreshdeskTickets_SEindia() {
  const FRESHDESK_DOMAIN = "graas-support.freshdesk.com";
  const API_KEY = "Gsiha9FK7J8YZBxJzR";
  const START_DATE = "2025-01-01T00:00:00Z";
  const encodedKey = Utilities.base64Encode(API_KEY + ":x");

  const baseUrl = `https://${FRESHDESK_DOMAIN}/api/v2/tickets?include=requester&per_page=100&updated_since=${encodeURIComponent(START_DATE)}`;

  let allTickets = [];
  let page = 1;

  while (true) {
    let url = `${baseUrl}&page=${page}`;
    let response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Basic " + encodedKey },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log("Error fetching tickets: " + response.getContentText());
      break;
    }

    let tickets = JSON.parse(response.getContentText());
    if (!tickets.length) break;

    let filteredTickets = tickets.filter(ticket =>
      ticket.custom_fields &&
      ticket.custom_fields.cf_client &&
      ticket.custom_fields.cf_client.trim().toLowerCase() === "seindia"
    );

    filteredTickets.forEach(ticket => {
      allTickets.push({
        type:ticket.custom_fields?.cf_issue_type || "",
        id: ticket.id,
        jira: ticket.custom_fields?.cf_jira_ticket_url || "",
        subject: ticket.subject || "",
        priority: "", // Skip reading priority
        status: getStatusText(ticket.status),
        task: ticket.type || "",
        createdAt: ticket.created_at || "",
        dueBy: ticket.due_by || "",
        updated_at: ticket.updated_at || "",
        requester: ticket.requester ? ticket.requester.name : "",
        assigneeId: getAgentName(ticket.responder_id),
        client: ticket.custom_fields?.cf_client || ""
      });
    });

    page++;
  }

  // Sort tickets by createdAt ascending (oldest first)
  allTickets.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets");
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Tickets");

  // Expanded headers: A..Q (existing) + R, S, T (placeholders) so Column T exists and is preserved
  let headers = [
    "Ticket type", "Ticket ID", "Jira Ticket", "Summary", "Priority", "Comments",
    "ETA", "Dependancy", "Updated ETA", "Status", "Task",
    "Created", "Due date", "Ticket Uodated Date", "Raised ticket by", "Graas Assignee", "Client",
    "Extra Col R", "Extra Col S", "Manual Override (Column T)"
  ];

  let currentData = sheet.getDataRange().getValues();
  if (currentData.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    currentData = [headers];
  }

  // Build map of existing rows keyed by Ticket ID (Column B)
  let sheetTicketMap = new Map();
  currentData.slice(1).forEach(row => {
    // ensure row has at least headers.length elements to avoid undefined indices
    while (row.length < headers.length) row.push("");
    sheetTicketMap.set(row[1], row); // Ticket ID is Column B
  });

  let updatedRows = [headers];

  allTickets.forEach(ticket => {
    if (sheetTicketMap.has(ticket.id)) {
      let oldRow = sheetTicketMap.get(ticket.id);

      // Manual override now comes from Column T (zero-based index 19)
      let manualStatus = oldRow[19];

      // If manual status exists, use it; otherwise use Freshdesk status
      let finalStatus = (manualStatus && manualStatus.toString().trim() !== "")
        ? manualStatus
        : ticket.status;

      updatedRows.push([
        ticket.type,              // A
        ticket.id,               // B
        ticket.jira,             // C
        ticket.subject,          // D
        oldRow[4],               // E - Priority (manual)
        oldRow[5],               // F - Comments (manual)
        oldRow[6],               // G - ETA (manual)
        oldRow[7],               // H - Dependancy (manual)
        oldRow[8],               // I - Updated ETA (manual)
        finalStatus,             // J - Status (computed)
        ticket.task,              // K - Task (kept from existing Task column)
        ticket.createdAt,        // L
        ticket.dueBy,            // M
        ticket.updated_at,       // N - Ticket Updated Date (from API)
        ticket.requester,        // O
        ticket.assigneeId,       // P
        ticket.client,           // Q
        oldRow[17] || "",        // R - preserved (if any)
        oldRow[18] || "",        // S - preserved (if any)
        oldRow[19] || ""         // T - Manual Override preserved
      ]);
    } else {
      // New ticket: create a row and include blanks for extra columns R,S,T
      updatedRows.push([
        ticket.cf_issue_type, ticket.id, ticket.jira, ticket.subject,
        "", "", "", "", "", ticket.status, "",   // A-K (J status, K Task blank)
        ticket.createdAt, ticket.dueBy, ticket.updated_at, // L-M-N
        ticket.requester, ticket.assigneeId, ticket.client, // O-P-Q
        "", "", "" // R, S, T blank for new rows
      ]);
    }
  });

  // Instead of clearing the sheet, preserve all columns not in columnsToUpdate
const existingData = sheet.getDataRange().getValues();

// Map header names to column indices
const colMap = {};
headers.forEach((h, idx) => colMap[h] = idx);

// Update only the columns you want to refresh
const columnsToUpdate = [
  "Type",
  "Task",
  "ID",
  "Jira",
  "Subject",
  "Status",
  "Created At",
  "Due By",
  "Updated At",
  "Requester",
  "Assignee",
  "Client",
  "Closed Date"
];

// Write full updatedRows back to the sheet (preserves manual columns already copied into updatedRows)
sheet.clearContents();
sheet.getRange(1, 1, updatedRows.length, headers.length).setValues(updatedRows);

Logger.log(`Updated ${allTickets.length} 'seindia' tickets`);

}
