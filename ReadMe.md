Google Sheets Unity Test
===

Requirements
---

* A project created in https://console.cloud.google.com/ with enabled Google Sheets API
* A Service Account created in https://console.cloud.google.com/iam-admin/serviceaccounts

Instructions
---

1. Share read access with the E-Mail of the Service Account for your spreadsheet
2. Create and download a key for the Service Account in JSON format
   * Name it `google-sheets-credentials.json`
   * Place it into `Assets/Resources` directory

Now you can run the project. In the top input fields enter the spreadsheet's id (can be copied from the URL), the sheet's name and the range of cells for data to retrieve.

The button `Read` retrieves the data in the simplest way, interpreting all the data as a plain text.

The button `Read More` retrieves the data, taking into account the data format which is set in Google Sheets for the cells.

In the large text field the result of reading the data is displayed, including values and their type, in JSON format.

---

Also you can watch [the video tutorial](https://www.youtube.com/watch?v=qm-Ooj6XjvE) I used while developing the project :-)