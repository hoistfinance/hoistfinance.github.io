// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function get_template_A_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += `<table style="color:#444; border: none; border-spacing: 0px;">`;
  str += `<tr>`;
  str += `  <td colspan="2" style="font-weight: bold">${user_info.name}`
  if (is_valid_data(user_info.greeting)) {
    str += ` <a style="font-size:10px;">${user_info.pronoun}</a>`;
  }
  str += `</td>`;
  str += `</tr>`;
  str += `<tr style="font-size: 14px">`;
  str += `  <td colspan="2"><a style="font-weight: bold">${user_info.job}</a>`; 
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
  str += `  <td colspan="2"><img style="width: 180px; padding-top: 5px;" src="../../../assets/full-logo.png" alt="Hoist Finance"/></td>`;
  str += `</tr>`;
  str += `<tr style="font-size: 12px">`;
  str += `  <td colspan="2"><a href="https://wwww.hoistfinance.com" style="color: #c71182;">www.hoistfinance.com</a></td>`;
  str += `</tr>`;
  str += `</table>`;

  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='padding-right: 5px;'><img src='../../../assets/hoist-finance.png' style='width: 60px;' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.job + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += `${user_info.greeting}<br/>`;
  }

  str += user_info.name;
  if(is_valid_data(user_info.pronoun)) {
    str += ` ${user_info.pronoun}`
  }

  str += '<br>___________________________________';

  if(is_valid_data(user_info.job)) {
    str += `<br>${user_info.job}`
  }

  if(is_valid_data(user_info.phone)) {
    str += `<br>${user_info.phone}`
  }
  
  return str;
}