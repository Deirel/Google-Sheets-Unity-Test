using System;
using System.IO;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using TMPro;
using UnityEngine;

public class SheetReader : MonoBehaviour
{
    [SerializeField] private TMP_InputField resultsInputField;
    [SerializeField] private TMP_InputField spreadsheetIdInputField;
    [SerializeField] private TMP_InputField sheetNameInputField;
    [SerializeField] private TMP_InputField dataRangeInputField;

    private SheetsService service;
    
    private string SpreadsheetId => spreadsheetIdInputField.text;

    private string SheetName => sheetNameInputField.text;
    
    private string DataRange => SheetName + "!" + dataRangeInputField.text;

    private void Awake()
    {
        const string credentialsFile = "google-sheets-credentials";

        var credentialsJson = Resources.Load<TextAsset>(credentialsFile).text;
        var stream = new MemoryStream(Encoding.UTF8.GetBytes(credentialsJson));
        var credentials = ServiceAccountCredential.FromServiceAccountData(stream);

        service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credentials
        });
    }

    public void Read() => TryRun(() =>
    {
        var request = service.Spreadsheets.Values.Get(SpreadsheetId, DataRange);
        var response = request.Execute();

        var data = response.Values.Select(row => row.Select(value => new
        {
            value,
            type = value?.GetType().Name
        }));

        var resultsJson = JsonConvert.SerializeObject(data, Formatting.Indented);
        resultsInputField.text = resultsJson;
    });

    public void ReadMore() => TryRun(() =>
    {
        var request = service.Spreadsheets.Get(SpreadsheetId);
        request.IncludeGridData = true;
        request.Ranges = new[] { DataRange };

        var response = request.Execute();
        var formats = response.Sheets
            // Find the sheet
            .First(s => s.Properties.Title == SheetName)
            // Get the grid data which corresponds to the `request.Ranges` above
            .Data.First()
            // Prepare for JSON
            .RowData.Select(rowData => rowData.Values.Select(cellData =>
            {
                var value = GetTypedCellValue(cellData);
                return new
                {
                    value, type = value?.GetType().Name
                };
            }));

        var resultsJson = JsonConvert.SerializeObject(formats, Formatting.Indented);
        resultsInputField.text = resultsJson;
    });

    private static object GetTypedCellValue(CellData cellData)
    {
        var effectiveValue = cellData.EffectiveValue;
        if (effectiveValue.ErrorValue != null)
            return effectiveValue.ErrorValue;

        if (effectiveValue.NumberValue != null)
            // NumberFormat may be null!
            return cellData.EffectiveFormat.NumberFormat?.Type switch
            {
                "DATE" or "DATE_TIME" => DateTime.FromOADate(effectiveValue.NumberValue.Value),
                "TIME" => TimeSpan.FromDays(effectiveValue.NumberValue.Value),
                _ => effectiveValue.NumberValue.Value
            };

        if (effectiveValue.BoolValue != null)
            return effectiveValue.BoolValue.Value;

        if (effectiveValue.StringValue != null)
            return effectiveValue.StringValue;

        return null;
    }
    
    private void TryRun(Action action)
    {
        try
        {
            action();
        }
        catch (Exception e)
        {
            Debug.LogError(e);
            resultsInputField.text = e.ToString();
        }
    }
}