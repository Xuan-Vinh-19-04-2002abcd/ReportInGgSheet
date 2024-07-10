using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Sheets.v4.Data;

public class Tests
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "Google Sheets API C#";
    static readonly string SpreadsheetId = "1mzdK6Ax_lHmxNTy-YzB1nFsC9ehqz9X5IOQ8rmjE8Ng";
    static readonly string SheetName = "ReportTest";
    static readonly string CredentialsPath = "C:\\Assignment HomeWork\\homework\\automationtest-429008-17f60284e21c.json";

    private IWebDriver driver;
    private SheetsService? service;

    [SetUp]
    public void Setup()
    {
        driver = new ChromeDriver();
        GoogleCredential credential;
        using (var stream = new FileStream(CredentialsPath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                .CreateScoped(Scopes);
        }
        
        service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });
    }

    [Test]
    [TestCase("https://gitlab.com/")]
    [TestCase("https://github.com/")]
    [TestCase("https://www.facebook.com/")]
    public void TestUrl(string url)
    {
        try
        {
            driver.Navigate().GoToUrl(url);
            string result = driver.Url.Contains(".com") ? "Pass" : "Fail";
            
            UpdateGoogleSheet(service, url, result, SheetName);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error testing URL {url}: {ex.Message}");
            UpdateGoogleSheet(service, url, "Fail", SheetName);
        }
    }

    [TearDown]
    public void TearDown()
    {
        driver.Quit();
        driver.Dispose();
        service.Dispose();
    }

    private void UpdateGoogleSheet(SheetsService service, string domain, string result, string sheetName)
    {
       
        var range = $"{sheetName}!A:B";
        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
        ValueRange response = request.Execute();
        IList<IList<object>> values = response.Values;
        
        int nextRow = values != null ? values.Count + 1 : 1;
        
        var valueRange = new ValueRange();
        valueRange.Values = new List<IList<object>> { new List<object> { domain, result } };
        
        var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, $"{sheetName}!A{nextRow}:B{nextRow}");
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

        var updateResponse = updateRequest.Execute();
    }
}
