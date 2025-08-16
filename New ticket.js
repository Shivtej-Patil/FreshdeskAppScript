function getStatusText(statusCode) {
  const statusMap = {
    2: "Open",
    3: "Pending",
    4: "Resolved",
    5: "Closed",
    8: "In Progress",
    12: "Clarification - Client/Third Party",
    13: "Tech Dependant",
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

function getLastResponsePreview(ticketId, encodedKey, domain) {
  try {
    let url = `https://${domain}/api/v2/tickets/${ticketId}/conversations`;
    let response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Basic " + encodedKey },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) return "";

    let conversations = JSON.parse(response.getContentText());
    if (!conversations.length) return "";

    let lastConv = conversations[conversations.length - 1];
    let text = (lastConv.body_text || "").replace(/\s+/g, " ").trim();

    return text.length > 150 ? text.substring(0, 150) + "..." : text;
  } catch (e) {
    Logger.log("Error fetching conversation for ticket " + ticketId + ": " + e);
    return "";
  }
}

function applyStatusColorFormatting(sheet, startRow, numRows) {
  const statusRange = sheet.getRange(startRow, 11, numRows, 1); // Column K

  // Clear existing background colors in the status column for the data rows
  statusRange.clearFormat();

  const backgrounds = [];

  const statuses = statusRange.getValues();

  statuses.forEach(row => {
    let color = ""; // default no color

    switch (row[0]) {
      case "Closed":
        color = "#008000"; // solid green
        break;
      case "Confirmation to Closure":
        color = "#90EE90"; // light green
        break;
      case "Clarification - Client/Third Party":
        color = "#FF4500"; // red/orange
        break;
      case "In Progress":
        color = "#FFA500"; // light orange
        break;
      case "Tech Dependant":
        color = "#800080"; // purple
        break;
      default:
        color = ""; // no color
    }
    backgrounds.push([color]);
  });

  statusRange.setBackgrounds(backgrounds);
}

function updateFreshdeskTickets_LandT() {
  const FRESHDESK_DOMAIN = "graas-support.freshdesk.com";
  const API_KEY = "freshdesk api";
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
      ticket.custom_fields.cf_client.trim().toLowerCase() === "l & t"
    );

    filteredTickets.forEach(ticket => {
      let lastResponse = getLastResponsePreview(ticket.id, encodedKey, FRESHDESK_DOMAIN);
      allTickets.push({
        type: ticket.type || "",
        id: ticket.id,
        jira: ticket.custom_fields?.cf_jira_ticket_url || "",
        subject: ticket.subject || "",
        priority: "", // Skip reading priority
        lastResponse: lastResponse,
        status: getStatusText(ticket.status),
        createdAt: ticket.created_at || "",
        dueBy: ticket.due_by || "",
        closedAt: ticket.closed_at || "",
        requester: ticket.requester ? ticket.requester.name : "",
        assigneeId: getAgentName(ticket.responder_id),
        client: ticket.custom_fields?.cf_client || ""
      });
    });

    page++;
  }

  // Sort tickets by createdAt ascending (oldest first)
  allTickets.sort((a, b) => {
    let dateA = new Date(a.createdAt);
    let dateB = new Date(b.createdAt);
    return dateA - dateB;
  });

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets");
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Tickets");

  let headers = [
    "Ticket type", "Ticket ID", "Jira Ticket", "Summary", "Priority", "Comments",
    "ETA", "Dependancy", "Updated ETA", "Last Response Preview", "Status", "Task",
    "Created", "Due date", "Ticket Closed Date", "Raised ticket by", "Graas Assignee", "Client"
  ];

  let currentData = sheet.getDataRange().getValues();
  if (currentData.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    currentData = [headers];
  }

  let sheetTicketMap = new Map();
  currentData.slice(1).forEach(row => {
    sheetTicketMap.set(row[1], row); // Ticket ID is Column B
  });

  let updatedRows = [headers];

  allTickets.forEach(ticket => {
    if (sheetTicketMap.has(ticket.id)) {
      let oldRow = sheetTicketMap.get(ticket.id);
      updatedRows.push([
        ticket.type,             // A
        ticket.id,               // B
        ticket.jira,             // C
        ticket.subject,          // D
        oldRow[4],               // E - Priority (manual)
        oldRow[5],               // F - Comments (manual)
        oldRow[6],               // G - ETA (manual)
        oldRow[7],               // H - Dependancy (manual)
        oldRow[8],               // I - Updated ETA (manual)
        ticket.lastResponse,     // J - Last Response Preview
        ticket.status,           // K
        oldRow[11],              // L - Task (manual)
        ticket.createdAt,        // M
        ticket.dueBy,            // N
        ticket.closedAt,         // O
        ticket.requester,        // P
        ticket.assigneeId,       // Q
        ticket.client            // R
      ]);
    } else {
      updatedRows.push([
        ticket.type, ticket.id, ticket.jira, ticket.subject,
        "", "", "", "", "", ticket.lastResponse, ticket.status, "",
        ticket.createdAt, ticket.dueBy, ticket.closedAt,
        ticket.requester, ticket.assigneeId, ticket.client
      ]);
    }
  });

  sheet.clearContents();
  sheet.getRange(1, 1, updatedRows.length, headers.length).setValues(updatedRows);

  // Apply status color formatting from row 2 to the last data row
  applyStatusColorFormatting(sheet, 2, updatedRows.length - 1);

  Logger.log(`Updated ${allTickets.length} 'L & T' tickets`);
}
