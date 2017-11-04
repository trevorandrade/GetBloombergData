using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
public static class DnaTestFunctions
{
    public static void Main()
    {
        string q = getPrice("7922:JP");
    }

    [ExcelFunction(Name = "getBloombergAttribute", Description = "Goes to https://www.bloomberg.com/quote/ticker where ticker is the first argument and grabs data based on xpath selector provided", Category = "My functions")]
    public static string getBloombergAttribute(
        [ExcelArgument(Name = "ticker", Description = "Bloomberg ticker e.g. 7922:JP")] string ticker,
        [ExcelArgument(Name = "selector", Description = "xpath selector which determine where to get the attribute on the webpage e.g. //h1[@class='name']")] string selector)
    {
        string url = generateBloombergURLfrom(ticker);
        return getAttribute(url, selector);
    }

    private static string generateBloombergURLfrom(string ticker)
    {
        return "https://www.bloomberg.com/quote/" + ticker;
    }


    // Bloomberg stores mcaps in a table and so we first have to search for text element 
    // which has market cap and then the actual value is in table element right after that
    [ExcelFunction(Name = "getMarketCap", Description = "Goes to https://www.bloomberg.com/quote/ticker and grabs marketcap", Category = "My functions")]
    public static double getMarketCap(
    [ExcelArgument(Name = "ticker", Description = "Bloomberg ticker e.g. 7922:JP")] string ticker)
    {
        string url = generateBloombergURLfrom(ticker);
        var web = new HtmlWeb();
        var doc = web.Load(url);
        string selector = "//*[contains(text(),'Market Cap')]";
        var foundElement = doc.DocumentNode.SelectSingleNode(selector); // find element with market cap
        // bloomberg site has Market Cap (b USD) and we try to extract b to know whether unit is bilions or millions
        string unit = Regex.Match(foundElement.InnerText, @"\((\w)").Groups[1].Value; 
        double multiplier = 1;
        switch (unit)
        {
            case "b": multiplier = 1E9; break;
            case "m": multiplier = 1E6; break;
        }
        // the actual value of market cap is the number immediately after the market cap string
        string marketCap = foundElement.NextSibling.NextSibling.InnerText;
        double marketCapValue = 1;
        Double.TryParse(marketCap, out marketCapValue); // convert mcap string to double
        return marketCapValue * multiplier;
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
