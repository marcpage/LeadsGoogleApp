// Services required: Sheets, People, Gmail

var sending_name = "Marc Page"
var sending_email = "Marc@ResolveToExcel.com"
var tracking_sheet_name = "Progress";

function get_tracking_sheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var tracking = spreadsheet.getSheetByName(tracking_sheet_name);
  if (!tracking) {
    tracking = spreadsheet.insertSheet(tracking_sheet_name, 0);
    tracking.getRange("A1:D1").setValues([["email", "Campaign", "Added Contact", "Replied"],]);
  }
  return tracking;
}

function get_row(tracking_data, campaign, email_address) {
  for (var row_index = 1; row_index < tracking_data.length; ++row_index) {
    if (tracking_data[row_index][0].length == 0) {
      return -1 * row_index;
    }
    if ((tracking_data[row_index][0] == email_address) 
        && (tracking_data[row_index][1] == campaign)) {
      return row_index;
    }
  }
  return 0;
}

function have_seen(tracking, campaign, email_address) {
  var actions = tracking.getRange("A2:D").getValues();
  for (var row_index = 0; row_index < actions.length; ++row_index) {
    if (actions[row_index][0].length == 0) {
      break;
    }
    if ((actions[row_index][0] == email_address) &&(actions[row_index][1] == campaign)) {
      return true;
    }
  }
  return false;
}

function have_emailed(tracking, campaign, email_address) {
  var actions = tracking.getRange("A2:D").getValues();
  for (var row_index = 0; row_index < actions.length; ++row_index) {
    if (actions[row_index][0].length == 0) {
      break;
    }
    if ((actions[row_index][0] == email_address) &&(actions[row_index][1] == campaign)) {
      return actions[row_index][3] == "Yes";
    }
  }
  return false;
}

function have_created_contact(tracking, campaign, email_address) {
  var actions = tracking.getRange("A2:D").getValues();
  for (var row_index = 0; row_index < actions.length; ++row_index) {
    if (actions[row_index][0].length == 0) {
      break;
    }
    if ((actions[row_index][0] == email_address) &&(actions[row_index][1] == campaign)) {
      return actions[row_index][2] == "Yes";
    }
  }
  return false;
}

function test_have_emailed() {
  var tracking = get_tracking_sheet();
  var emailed = have_emailed(tracking, "Form Responses 1", "marc@resolvetoexcel.com");
  var contact = have_created_contact(tracking, "Form Responses 1", "marc@resolvetoexcel.com");
  console.log([emailed, contact]);
}

function get_every_group() {
  var groups = People.ContactGroups.list();
  var group_mappings = {};
  // TODO: Use pageToken and nextPageToken for more than 100 groups
  for (var index = 0; index < groups.contactGroups.length; ++index) {
    group = groups.contactGroups[index];
    group_mappings[group.resourceName] = group.name;
  }
  return group_mappings;
}

function test_get_every_group() {
  console.log(get_every_group());
}

function get_every_contact() {
  var groups = get_every_group();
  var fields = {"personFields": "names,emailAddresses,memberships,biographies"};
  var everyone = {};
  var people; 
  do {
    people = People.People.Connections.list("people/me", fields);
    for (var person_index=0; person_index < people.connections.length; ++person_index) {
      person = people.connections[person_index];
      if (!person.names) {
        continue;
      }
      notes = person.biographies;
      emails = person.emailAddresses 
                ? person.emailAddresses.map(function (address) {return address.value}) 
                : [];
      membership = person.memberships 
                ? person.memberships.map(function (group) {
                    return groups[group.contactGroupMembership.contactGroupResourceName]}) 
                : [];
      everyone[person.resourceName] = {
          "first": person.names[0].givenName,
          "last": person.names[0].familyName,
          "emails": emails,
          "groups": membership,
          "notes": "",
      }
      for (var note_index = 0; note_index < (notes ? notes.length : 0); ++ note_index) {
        everyone[person.resourceName]["notes"] += notes[note_index].value;
      }
    }
    fields.pageToken = people.nextPageToken;
  } while(people.nextPageToken);
  return everyone;
}

function test_get_every_contact() {
  var everyone = get_every_contact();
  console.log(everyone);
}

function match_email_attechments_from_body(inline_images, body) {
  var name_image = {};
  var inline_pattern = /<img[^>]+data-surl="cid:([^"]+)" src="cid:([^"]+)" alt="([^"]+)"[^>]+>/g;
  var inlined_image_info = [...body.matchAll(inline_pattern)];
  for (var inlined_index = 0; inlined_index < inlined_image_info.length; ++inlined_index) {
    if (inlined_image_info[inlined_index][1] != inlined_image_info[inlined_index][2]) {
      throw "Mismatched identifiers";
    }
  }
  if (inlined_image_info.length != inline_images.length) {
    throw "Mismatched inlined image counts";
  }
  for (var images_index = 0; images_index < inline_images.length; ++images_index) {
    var image = inline_images[images_index];
    var image_info = inlined_image_info[images_index];
    if (image_info[3] != image.getName()) {
      throw "Image name mismatch: " + image_info[3] + " vs " + image.getName();
    }
    name_image[image_info[1]] = image;
  }
  return name_image;
}

function replace_fields(text, fields) {
  var variable_pattern = /\{\{([A-Za-z0-9]+)\}\}/g;
  var replacements = {};
  var instances = [...text.matchAll(variable_pattern)];
  for (var index = 0; index < instances.length; ++index) {
    var key = instances[index][0];
    var value = fields[instances[index][1]];
    replacements[key] = value ? value : ""; 
  }
  for (var key in replacements) {
    text = text.replace(key, replacements[key]);
  }
  return text;
}

function send_message(draft, email, fields) {
  var inline_images = draft.getAttachments({
      "includeAttachments": false, 
      "includeInlineImages": true});
  var attachments = draft.getAttachments({
      "includeAttachments": true, 
      "includeInlineImages": false});
  var plain_body = draft.getPlainBody();
  var html_body = draft.getBody();
  GmailApp.sendEmail(email, draft.getSubject(), replace_fields(plain_body, fields),
      {
        "htmlBody": replace_fields(html_body, fields),
        "attachments": attachments,
        "cc": draft.getCc(),
        "bcc": draft.getBcc(),
        "replyTo": sending_name + " <" + sending_email + ">",
        "name": sending_name,
        "inlineImages": match_email_attechments_from_body(inline_images, draft.getBody())
      }
    );
}

function send_draft(subject, email, fields) {
  var drafts = GmailApp.getDraftMessages();
  for (var index = 0; index < drafts.length; ++index) {
    if (drafts[index].getSubject() == subject) {
      console.log(drafts[index].getSubject());
      send_message(drafts[index], email, fields);
      break;
    }
  }
}

function add_new_submissions_from_form(form_sheet, tracking) {
  var campaign = form_sheet.getSheetName();
  var form_data = form_sheet.getRange("A1:Z").getValues();
  var email_column = -1;
  for (var column = 0; column < form_data[0].length; ++column) {
    if (form_data[0][column].toLowerCase().indexOf("email") >= 0) {
      email_column = column;
      break;
    }
  }
  if (-1 == email_column) {
    console.log(campaign + " does not appear to have an email column");
  }
  var tracking_range = tracking.getRange("A1:B");
  var tracking_data = tracking_range.getValues();
  var values_added = false;
  for (var row = 1; row < form_data.length; ++row) {
    if (form_data[row][0].length == 0) {
      break; // stop once we reach an empty first cell (timestamp)
    }
    var email_address = form_data[row][email_column];
    var tracking_row = get_row(tracking_data, campaign, email_address);
    if (tracking_row >= 0) {
      continue;
    }
    console.log(tracking_row);
    tracking_data[-1 * tracking_row][0] = email_address;
    tracking_data[-1 * tracking_row][1] = campaign;
    values_added = true;
  }
  if (values_added) {
    tracking_range.setValues(tracking_data);
  }
}

function send_pending_emails(tracking) {
  var log_range = tracking.getRange("A1:D");
  var log = log_range.getValues();
  for (var row = 0; row < log.length; ++row) {
    if (log[row][0].length == 0) {
      break; // stop at first row with first column empty
    }
    if (log[row][3].length == 0) {
      var email_address = log[row][0];
      var campaign = log[row][1];
      try {
        send_draft(campaign, email_address, {});
        log[row][3] = "Sent";
        log_range.setValues(log);
      } catch {
        console.log("Unable to send email to " + email_address + " for campaign " + campaign);
      }
      
    }
  }
}

function handle_new_form_submissions() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();
  var tracking = get_tracking_sheet();
  for (var i = 0; i < sheets.length; ++i) {
    var sheet = sheets[i];
    if (sheet.getFormUrl() == null) {
      continue;
    }
    add_new_submissions_from_form(sheet, tracking);
  }
  send_pending_emails(tracking);
}
