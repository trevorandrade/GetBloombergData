using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using HtmlAgilityPack;
public static class DnaTestFunctions
{
    [ExcelFunction(Name = "getBloombergAttribute", Description = "Goes to https://www.bloomberg.com/quote/ticker where ticker is the first argument and grabs data based on xpath selector provided", Category = "My functions")]
    public static string getBloombergAttribute(
        [ExcelArgument(Name = "ticker", Description = "Bloomberg ticker e.g. 7922:JP")] string ticker,
        [ExcelArgument(Name = "selector", Description = "xpath selector which determine where to get the attribute on the webpage e.g. //h1[@class='name']")] string selector)
    {
        string url = "https://www.bloomberg.com/quote/" + ticker;
        return getAttribute(url, selector);
    }

    [ExcelFunction(Name = "getAttribute", Description = "Get an attribute on a webpage using xpath expression", Category = "My functions")]
    public static string getAttribute(
               [ExcelArgument(Name = "url", Description = "webpage to get data from")] string url,
               [ExcelArgument(Name = "selector", Description = "xpath expression to select element of page to grab data from")] string selector)
    {
        // From Web
        var web = new HtmlWeb();
        var doc = web.Load(url);
        var val = doc.DocumentNode.SelectSingleNode(selector).InnerText;
        return val.Trim();
    }

    [ExcelFunction(Name = "getName", Description = "Get name of stock from Bloomberg webpage using ticker", Category = "My functions")]
    public static string getName(
        [ExcelArgument(Name = "ticker", Description = "BLoomberg ticker e.g. 7922:JP")] string ticker)
    {
        return getBloombergAttribute(ticker, "//h1[@class='name']");
    }

    [ExcelFunction(Name = "getPrice", Description = "Get price of stock from Bloomberg webpage using ticker", Category = "My functions")]
    public static string getPrice(
        [ExcelArgument(Name = "ticker", Description = "BLoomberg ticker e.g. 7922:JP")] string ticker)
    {
        return getBloombergAttribute(ticker, "//div[@class='price']");
    }
}
