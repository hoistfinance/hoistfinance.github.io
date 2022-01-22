// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on both Outlook on web and Outlook on Windows.

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows only. Insert signature into appointment or message.
 * Outlook on Windows can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the 'Email Signature' add-in.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "hoist-finance.png";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += `<table style="color:#444; border: none; border-spacing: 0px;">`;
  str += `<tr>`;
  str += `  <td colspan=2 style="font-weight: bold">${user_info.name}</td>`;
  str += `</tr>`;
  str += `<tr style="font-size: 14px">`;
  str += `  <td colspan=2><a style="font-weight: bold">${user_info.job}</a>`; 
  if (is_valid_data(user_info.department)) {
    str += ` | ${user_info.department}` ;
  }
  str += `</td>`;
  str += `</tr>`;
  str += `<tr style="font-size: 13px">`;
  str += `  <td style="width: 35px">Email</td>`;
  str += `  <td>: ${user_info.email}</td>`;
  str += `</tr>`;
  str += `<tr style="font-size: 13px">`;
  str += `  <td style="width: 35px">Phone</td>`;
  str += `  <td>: ${user_info.phone}</td>`;
  str += `</tr>`;
  str += `<tr>`;
  str += `  <td><img style="width: 180px; padding-top: 5px;" src="cid:${logoFileName}" alt="Hoist Finance"/></td>`;
  str += `</tr>`;
  str += `<tr style="font-size: 12px">`;
  str += `  <td colspan=2><a href="https://wwww.hoistfinance.com" style="color: #c71182;">www.hoistfinance.com</a></td>`;
  str += `</tr>`;
  str += `</table>`;

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
    "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",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += `${user_info.greeting}<br/>`;
  }

  str += user_info.name;

  if(is_valid_data(user_info.job)) {
    str += `<br>${user_info.job}`
  }

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
