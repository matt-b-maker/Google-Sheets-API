using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using HtmlAgilityPack;
using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// This script uses three Nuget packages:
/// 1) Google Sheets API, requires a JSON OAuth key and authorization from the admin of the sheet it's trying to access
/// 2) Puppeteer Sharp to get the fully loaded html from the Rocket League url
/// 3) Html Agility Pack to parse through the html retrieved by Puppeteer
/// </summary>

namespace GoogleSheetsAPITest
{
    class Player
    {
        public Player (List<string> websites, int rowNumber)
        {
            Websites = websites;
            RowNumber = rowNumber;

            OnesMMR = 0;
            TwosMMR = 0;
            ThreesMMR = 0;
        }
        public List<string> Websites { get; set; }
        public int RowNumber { get; set; }

        public int OnesMMR { get; set; }
        public int TwosMMR { get; set; }
        public int ThreesMMR { get; set; }
    }

    class Playlist
    {
        public Playlist(string type)
        {
            Type = type;
        }
        public string Type { get; set; }
        public string MMR { get; set; }
    }

    class Program
    {
        static Dictionary<int, string> tabSelections = new Dictionary<int, string>()
        {
            { 1 , "Add a sheet name" },
            { 2 , "Remove a sheet name" }
        };

        //Starting variables for the Google API
        static string[] scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Yeeeeeee HAWWWWWWWW!!!!!!!";
        static readonly string SpreadhseetId = "13b3_3VDtnf2nhbMbMJFSzsKK4cGQjHJcIuGYRp6rYD8";
        static string sheet = "";
        static string sheetRange;
        static SheetsService service;

        const string PlaylistNodeSpec = "//div[@class='playlist']";
        const string MMRNodeSpec = "//div[@class='mmr']";

        const string onesColHeader = "1's";
        static char onesCol, twosCol, threesCol;

        //Need to modify this list in different areas of the script, so declaring here
        public static List<Player> players = new List<Player>();

        static async Task Main()
        {
            PopulateTabSelectionDict();

            DisplayOptions();

            Console.WriteLine(sheet);

            sheetRange = $"{sheet}!A1:Z1000";

            //Obviously, connect to the google sheet first using the Google Sheets API package
            ConnectToSheet();

            //Read the info from the sheet and populate the list of Player objects using the Google Sheets API package
            try
            {
                ReadEntries(sheetRange, onesColHeader);
            }
            catch
            {
                Console.WriteLine($"{sheet} is not a valid sheet name, punk-ass. I'm removing it from the list. Or something else went wrong, I don't know.");
                foreach (var tab in tabSelections)
                {
                    if (tab.Value == sheet)
                    {
                        RemoveSheetFromDictWithKey(tabSelections, tab.Key);
                        break;
                    }
                }
                Console.WriteLine("Good luck in life, kiddo. Relaunch this shit and try again.");
                Console.ReadKey();
                return;
            }

            try
            {
                await ScrapeAndUpdate();
            }
            catch (Exception e)
            {
                Console.WriteLine($"\n\n{e.Message}");
            }
        }

        private static void DisplayOptions()
        {
            List<string> officialLines = new List<string>();
            Console.Clear();
            for (int i = 1; i <= tabSelections.Count; i++)
            {
                Console.WriteLine($"{i}: {tabSelections[i]}");
            }

            do
            {
                sheet = GetSheetName(tabSelections);
                if (sheet == "Add a sheet name")
                {
                    Console.Clear();
                    Console.WriteLine("Enter the name of the sheet you want to add to the list: ");
                    string newItem = Console.ReadLine();
                    tabSelections.Add(tabSelections.Count + 1, newItem);
                    string[] lines = File.ReadAllLines("Sheets for Matt.txt");
                    foreach (var line in lines)
                    {
                        officialLines.Add(line);
                    }
                    officialLines.Add(newItem);
                    File.WriteAllLines("Sheets for Matt.txt", officialLines);

                    officialLines.Clear();

                    Console.Clear();

                    for (int i = 1; i <= tabSelections.Count; i++)
                    {
                        Console.WriteLine($"{i}: {tabSelections[i]}");
                    }
                }
                else if (sheet == "Remove a sheet name")
                {
                    Console.WriteLine("\nEnter the number of the sheet you want to remove: ");
                    RemoveSheetFromDict(tabSelections);
                    Console.Clear();

                    for (int i = 1; i <= tabSelections.Count; i++)
                    {
                        Console.WriteLine($"{i}: {tabSelections[i]}");
                    }
                }
            } while (sheet == "" || sheet == "Add a sheet name" || sheet == "Remove a sheet name");
        }

        private static void PopulateTabSelectionDict()
        {
            string tabSelectionsTextPre = File.ReadAllText("Sheets for Matt.txt");
            string[] tabSelectionsText = tabSelectionsTextPre.Split("\r\n");
            for (int i = 1; i <= tabSelectionsText.Length; i++)
            {
                if (tabSelectionsText[i - 1] != "") 
                {
                    tabSelections.Add(i + 2, tabSelectionsText[i - 1]);
                }
            }
        }

        private static async Task ScrapeAndUpdate()
        {
            //Use Puppeteer Sharp to open a single browser before going to all the different web pages. MEMORY MANAGEMENT, DAAAAAAAWG
            var options = new LaunchOptions
            {
                Headless = true,
                DefaultViewport = null,
                ExecutablePath = @"C:\Program Files (x86)\Google\Chrome\Application"
            };

            var browser = await Puppeteer.LaunchAsync(options);

            //await new BrowserFetcher().DownloadAsync();

            //var browser = await Puppeteer.LaunchAsync(new LaunchOptions
            //{
            //    Headless = true,
            //    DefaultViewport = null
            //});

            List<Playlist> playlists = new List<Playlist>();

            if (players.Count > 0)
            {
                foreach (var player in players)
                {
                    if (player.Websites.Count > 0)
                    {
                        //only do the thing once with no comparisons if there is only one website
                        if (player.Websites.Count == 1)
                        {
                            //This is in a try/catch to avoid a possible failure after an exception is thrown in the GetHtml() method
                            try
                            {
                                //Use Puppeteer Sharp in the GetHtml() method to retrieve the html from the website in the player object
                                string content = await GetHtml(player.Websites[0], browser);
                                HtmlDocument doc = new HtmlDocument();
                                doc.LoadHtml(content);

                                //Use html agility pack to get every node associated with div elements of classes "playlist" and "mmr"
                                var playList = doc.DocumentNode.SelectNodes(PlaylistNodeSpec);
                                var mmrList = doc.DocumentNode.SelectNodes(MMRNodeSpec);

                                //Populate the list of Playlist objects with the type first, then the mmr
                                if (playList != null)
                                {
                                    foreach (var playlist in playList)
                                    {
                                        playlists.Add(new Playlist(playlist.InnerText));
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("playlists returned null");
                                    Console.ReadKey();
                                }

                                if (mmrList != null)
                                {
                                    for (int i = 0; i < mmrList.Count; i++)
                                    {
                                        playlists[i].MMR = mmrList[i].InnerText;
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("mmrlists returned null");
                                    Console.ReadKey();
                                }

                                //Update player object mmr values
                                foreach (var playlist in playlists)
                                {
                                    if (playlist.Type.Contains("Ranked Duel 1v1 "))
                                        player.OnesMMR = ConvertToInt(playlist.MMR);
                                    else if (playlist.Type.Contains("Ranked Doubles 2v2 "))
                                        player.TwosMMR = ConvertToInt(playlist.MMR);
                                    else if (playlist.Type.Contains("Ranked Standard 3v3 "))
                                        player.ThreesMMR = ConvertToInt(playlist.MMR);
                                    else
                                        continue;
                                }
                            }
                            catch (Exception e)
                            {
                                //Head to the next iteration in the loop if GetHtml() throws a timeout exception
                                Console.WriteLine(e.Message);
                                CreateEntry(player.RowNumber, "------------------", onesCol);
                                CreateEntry(player.RowNumber, "error with website", twosCol);
                                CreateEntry(player.RowNumber, "------------------", threesCol);
                                continue;
                            }
                        }
                        //Gonna have to compare values from each website and find the highest value among them
                        //if there is more than one website to look through
                        else
                        {
                            foreach (var website in player.Websites)
                            {
                                try
                                {
                                    string content = await GetHtml(website, browser);
                                    HtmlDocument doc = new HtmlDocument();
                                    doc.LoadHtml(content);

                                    var playList = doc.DocumentNode.SelectNodes(PlaylistNodeSpec);
                                    var mmrList = doc.DocumentNode.SelectNodes(MMRNodeSpec);

                                    if (playList != null)
                                    {
                                        foreach (var playlist in playList)
                                        {
                                            playlists.Add(new Playlist(playlist.InnerText));
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("playlists returned null");
                                        Console.ReadKey();
                                    }

                                    if (mmrList != null)
                                    {
                                        for (int i = 0; i < mmrList.Count; i++)
                                        {
                                            playlists[i].MMR = mmrList[i].InnerText;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("mmrlists returned null");
                                        Console.ReadKey();
                                    }

                                    foreach (var playlist in playlists)
                                    {
                                        if (playlist.Type.Contains("Ranked Duel 1v1 "))
                                        {
                                            if (ConvertToInt(playlist.MMR) > player.OnesMMR)
                                            {
                                                player.OnesMMR = ConvertToInt(playlist.MMR);
                                            }
                                        }
                                        else if (playlist.Type.Contains("Ranked Doubles 2v2 "))
                                        {
                                            if (ConvertToInt(playlist.MMR) > player.TwosMMR)
                                            {
                                                player.TwosMMR = ConvertToInt(playlist.MMR);
                                            }
                                        }
                                        else if (playlist.Type.Contains("Ranked Standard 3v3 "))
                                        {
                                            if (ConvertToInt(playlist.MMR) > player.ThreesMMR)
                                            {
                                                player.ThreesMMR = ConvertToInt(playlist.MMR);
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                    playlists.Clear();
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    CreateEntry(player.RowNumber, "------------------", onesCol);
                                    CreateEntry(player.RowNumber, "error with website", twosCol);
                                    CreateEntry(player.RowNumber, "------------------", threesCol);
                                    continue;
                                }
                            }
                        }

                        CreateEntry(player.RowNumber, player.OnesMMR.ToString(), onesCol);
                        CreateEntry(player.RowNumber, player.TwosMMR.ToString(), twosCol);
                        CreateEntry(player.RowNumber, player.ThreesMMR.ToString(), threesCol);
                        playlists.Clear();
                    }
                }
            }

            //Close the browser once everything is done
            await browser.CloseAsync();
        }

        private static void RemoveSheetFromDict(Dictionary<int, string> tabSelections)
        {
            ConsoleKeyInfo input = Console.ReadKey();
            int key;

            List<string> values = new List<string>();

            if (char.IsDigit(input.KeyChar))
            {
                key = int.Parse(input.KeyChar.ToString());
                if (tabSelections.ContainsKey(key) && key != 1 && key != 2)
                {
                    List<string> lines = new List<string>();
                    string[] fileLines = File.ReadAllLines("Sheets for Matt.txt");
                    foreach (var line in fileLines)
                        lines.Add(line);

                    for (int i = 0; i < lines.Count; i++)
                    {
                        if (lines[i] == tabSelections[key])
                        {
                            lines.Remove(lines[i]);
                            break;
                        }
                    }

                    File.WriteAllLines("Sheets for Matt.txt", lines);
                    tabSelections.Remove(key);
                    foreach (var tab in tabSelections)
                    {
                        if (tab.Key != 1 && tab.Key != 2)
                        {
                            values.Add(tab.Value);
                        }
                    }
                    tabSelections.Clear();
                    tabSelections.Add(1, "Add a sheet name");
                    tabSelections.Add(2, "Remove a sheet name");
                    for (int i = 0; i < values.Count; i++)
                    {
                        tabSelections.Add(i + 3, values[i]);
                    }
                }
                else
                    Console.WriteLine("Nope");
            }
        }

        private static void RemoveSheetFromDictWithKey(Dictionary<int, string> tabSelections, int key)
        {
            List<string> values = new List<string>();

            if (tabSelections.ContainsKey(key) && key != 1 && key != 2)
            {
                List<string> lines = new List<string>();
                string[] fileLines = File.ReadAllLines("Sheets for Matt.txt");
                foreach (var line in fileLines)
                    lines.Add(line);

                for (int i = 0; i < lines.Count; i++)
                {
                    if (lines[i] == tabSelections[key])
                    {
                        lines.Remove(lines[i]);
                        break;
                    }
                }

                File.WriteAllLines("Sheets for Matt.txt", lines);
                tabSelections.Remove(key);
                foreach (var tab in tabSelections)
                {
                    if (tab.Key != 1 && tab.Key != 2)
                    {
                        values.Add(tab.Value);
                    }
                }
                tabSelections.Clear();
                tabSelections.Add(1, "Add a sheet name");
                tabSelections.Add(2, "Remove a sheet name");
                for (int i = 0; i < values.Count; i++)
                {
                    tabSelections.Add(i + 3, values[i]);
                }
            }
        }

        private static string GetSheetName(Dictionary<int, string> tabSelections)
        {
            ConsoleKeyInfo input = Console.ReadKey();
            int key;

            if (char.IsDigit(input.KeyChar))
            {
                key = int.Parse(input.KeyChar.ToString());
                if (tabSelections.ContainsKey(key))
                    return tabSelections[key];
                else
                    Console.CursorLeft = 0;
                    Console.WriteLine("Nope");
                    return "";
            }

            return "";
        }

        private static int ConvertToInt(string mMR)
        {
            return Convert.ToInt32(mMR.Replace(",", ""));
        }

        private static async Task<string> GetHtml(string website, Browser browser)
        {
            Stopwatch stopwatch = new Stopwatch();

            Console.WriteLine($"Running scrape on {website}");
            string url = website;
            string content;

            using (var page = await browser.NewPageAsync())
            {
                stopwatch.Start();
                WaitForSelectorOptions timeout = new WaitForSelectorOptions();
                timeout.Timeout = 10000;

                await page.GoToAsync(url, WaitUntilNavigation.DOMContentLoaded);
                await page.WaitForSelectorAsync("div.mmr", timeout);
                content = await page.GetContentAsync();
                await page.CloseAsync();
                Console.WriteLine($"It took {stopwatch.ElapsedMilliseconds} ms to scrape");
                stopwatch.Stop();
            }

            return content;
        }

        private static void ConnectToSheet()
        {
            UserCredential credential;

            using (var stream = new FileStream("SHAclient_secret.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None)//,
                                           //new FileDataStore(credPath, true))
                    .Result;

                Console.WriteLine("Credential file saved to: " + credPath);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        static void ReadEntries(string range, string ones)
        {
            var request = service.Spreadsheets.Values.Get(SpreadhseetId, range);

            var response = request.Execute();
            var values = response.Values;

            GetColumns(values, ones);

            int rowNumber = 2;

            if (values != null && values.Count > 0)
            {
                foreach (var value in values)
                {
                    if (value != null && value != values[0])
                    {
                        players.Add(new Player(GetWebsitesForEachPlayer(value), rowNumber));
                        rowNumber++;
                    }
                }
            }
            else
            {
                Console.WriteLine("Didn't get any data");
                return;
            }

            if (players.Count > 0)
            {
                foreach (var player in players)
                {
                    Console.WriteLine(player.RowNumber);
                    foreach (var website in player.Websites)
                        Console.WriteLine(website);
                }
            }
        }

        private static void GetColumns(IList<IList<object>> values, string ones)
        {
            int letterCount = 65;
            for (int i = 0; i < values.Count; i++)
            {
                if (values[0][i].ToString().Contains(ones))
                {
                    onesCol = Convert.ToChar(letterCount);
                    twosCol = Convert.ToChar(letterCount + 1);
                    threesCol = Convert.ToChar(letterCount + 2);
                    break;
                }
                else
                {
                    letterCount++;
                }
            }
        }

        private static List<string> GetWebsitesForEachPlayer(IList<object> value)
        {
            List<string> websites = new List<string>();
            foreach (var cell in value)
            {
                if (cell.ToString().Contains("https://"))
                {
                    websites.Add(cell.ToString());
                }
            }
            return websites;
        }

        static void CreateEntry(int rowNumber, string mmr, char column)
        {
            var range = $"{sheet}!{column}{rowNumber}:{column}{rowNumber}";
            var valueRange = new ValueRange();

            if (mmr == "0")
                mmr = "Not Ranked";
            else
                mmr = AddCommaIfNecessary(mmr);

            var objectList = new List<object>() { mmr };
            valueRange.Values = new List<IList<object>> { objectList };

            var enterInfo = service.Spreadsheets.Values.Update(valueRange, SpreadhseetId, range);
            enterInfo.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var infoEntered = enterInfo.Execute();
        }

        private static string AddCommaIfNecessary(string mmr)
        {
            if (mmr.Length == 4)
            {
                char newMMR_A = mmr[0];
                string newMMR_B = $"{mmr[1]}{mmr[2]}{mmr[3]}";
                return newMMR_A + "," + newMMR_B;
            }
            else
            {
                return mmr;
            }
        }
    }
}
