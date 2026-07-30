using HtmlAgilityPack;
using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace BizhubReader
{
    enum PrinterType
    {
        Offline,
        Legacy,
        Spa
    }

    class CounterInfo
    {
        public int PrintBlack { get; set; }
        public int PrintColor { get; set; }
        public int Scans { get; set; }
        public string SerialNumber { get; set; } = "";
    }

    class PreviousCounter
    {
        public int Black { get; set; }
        public int Color { get; set; }
        public int Scan { get; set; }
    }

    class PrinterResult
    {
        public string Name { get; set; } = "";       
        public string EszkozKod { get; set; } = "";
        public string Serial { get; set; } = "";     
        public string IP { get; set; } = "";
        public string Type { get; set; } = "";
        public string Cim { get; set; } = "";
        public CounterInfo? Counter { get; set; }
    }

    class Printer
    {
        public string IP { get; set; } = "";
    }


    class Program
    {
        static readonly CookieContainer cookies = new CookieContainer();

        static readonly HttpClientHandler handler = new HttpClientHandler
        {
            CookieContainer = cookies,
            UseCookies = true
        };

        static readonly HttpClient client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(30)
        };


        static async Task Main(string[] args)
        {
            var printers = LoadPrinters("ips.txt");

            Dictionary<string, PreviousCounter> previous = LoadPreviousCounters();

            List<PrinterResult> results = new();

            foreach (var ip in printers)
            {
                string gyariSzam = "";

                Console.WriteLine("-------------------------");
                Console.WriteLine($"IP: {ip}");

                PrinterType type = await DetectPrinterType(ip);

                switch (type)
                {
                    case PrinterType.Legacy:
                        Console.WriteLine("Típus: régi Bizhub");

                        gyariSzam = await ReadLegacySerial(ip) ?? "";

                        CounterInfo? counter = await ReadLegacyPrinter(ip);

                        Console.WriteLine($"Serial: {gyariSzam}");
                        Console.WriteLine($"Black : {counter.PrintBlack}");
                        Console.WriteLine($"Color : {counter.PrintColor}");
                        Console.WriteLine($"Scans : {counter.Scans}");

                        if (counter != null)
                        {
                            results.Add(new PrinterResult
                            {
                                Serial = gyariSzam,
                                IP = ip,
                                Type = "Régi Bizhub",
                                Counter = counter
                            });
                        }

                        break;

                    case PrinterType.Spa:
                        Console.WriteLine("Típus: új SPA Bizhub");

                        CounterInfo? spaCounter = await ReadSpaPrinter(ip);

                        if (spaCounter != null)
                        {
                            gyariSzam = spaCounter.SerialNumber;

                            Console.WriteLine($"Serial: {gyariSzam}");
                            Console.WriteLine($"Black : {spaCounter.PrintBlack}");
                            Console.WriteLine($"Color : {spaCounter.PrintColor}");
                            Console.WriteLine($"Scans : {spaCounter.Scans}");

                            results.Add(new PrinterResult
                            {
                                Serial = gyariSzam,
                                IP = ip,
                                Type = "SPA Bizhub",
                                Counter = spaCounter
                            });
                        }

                        break;

                    default:
                        Console.WriteLine(string.IsNullOrEmpty(gyariSzam)
            ? "Sorozatszám: nem elérhető"
            : $"Sorozatszám: {gyariSzam}");

                        results.Add(new PrinterResult
                        {
                            Serial = gyariSzam,
                            IP = ip,
                            Type = "Offline",
                            Counter = null
                        });

                        break;
                }
            }

            SaveResultsExcel(results, previous);

            Console.WriteLine("Kész");
        }

            static List<string> LoadPrinters(string file)
            {
                var ips = new List<string>();

                if (!File.Exists(file))
                {
                    Console.WriteLine($"Nem található: {file}");
                    return ips;
                }

                foreach (string line in File.ReadAllLines(file))
                {
                    string trimmed = line.Trim();

                    if (string.IsNullOrWhiteSpace(trimmed))
                        continue;

                    if (trimmed.StartsWith("#"))
                        continue;

                    ips.Add(trimmed);
                }

                Console.WriteLine($"Betöltött IP-k: {ips.Count} db");

                return ips;
            }


        //Megállapítja, hogy az adott IP címen milyen típusú Konica Minolta nyomtató található. (XML/SPA/Offline)
        static async Task<PrinterType> DetectPrinterType(string ip)
        {
            if (await UrlExists($"http://{ip}/wcd/system_counter.xml"))
                return PrinterType.Legacy;


            if (await UrlExists($"http://{ip}/wcd/spa_main.html"))
                return PrinterType.Spa;


            return PrinterType.Offline;
        }


        //URL elérhető e
        static async Task<bool> UrlExists(string url)
        {
            try
            {
                using var response =
                    await client.GetAsync(url);

                return response.StatusCode == HttpStatusCode.OK;
            }
            catch
            {
                return false;
            }
        }


        //Régi(XML) nyomtatók kiolvasása
        static async Task<CounterInfo?> ReadLegacyPrinter(string ip)
        {
            try
            {
                await client.GetAsync($"http://{ip}/wcd/index.html?access=SYS_COU");

                string url = $"http://{ip}/wcd/system_counter.xml";

                using var request = new HttpRequestMessage(HttpMethod.Post, url);

                request.Content = new StringContent(
                    "usr=S_COU;",
                    Encoding.UTF8,
                    "text/plain"
                );

                var response = await client.SendAsync(request);

                string text = await response.Content.ReadAsStringAsync();

                return ParseCounter(text);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return null;
            }
        }

        static async Task<string?> ReadLegacySerial(string ip)
        {
            try
            {
                await client.GetAsync($"http://{ip}/wcd/index.html?access=SYS_COU");

                string url = $"http://{ip}/wcd/system_device.xml";

                string xml = await client.GetStringAsync(url);

                var doc = XDocument.Parse(xml);

                return doc.Descendants("SerialNumber").FirstOrDefault()?.Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Serial olvasási hiba ({ip}): {ex.Message}");
                return null;
            }
        }

        static CounterInfo ParseCounter(string xml)
        {
            var doc = XDocument.Parse(xml);

            int copyBlack = GetValue(doc, "CopyCounter", "BwTotal");
            int printBlack = GetValue(doc, "PrintCounter", "BwTotal");

            int copyColor = GetValue(doc, "CopyCounter", "FullColorTotal");
            int printColor = GetValue(doc, "PrintCounter", "FullColorTotal");
            int biColor = GetValue(doc, "PrintCounter", "BiColorTotal");

            int scans = GetValue(doc, "ScanFaxCounter", "DocumentReadTotal");

            return new CounterInfo
            {
                PrintBlack = copyBlack + printBlack,
                PrintColor = copyColor + printColor + biColor,
                Scans = scans
            };
        }

        static string? FindSerialNumber(JsonElement element)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (var property in element.EnumerateObject())
                {
                    if (property.Name == "SerialNumber" && property.Value.ValueKind == JsonValueKind.String)
                    {
                        return property.Value.GetString();
                    }

                    var found = FindSerialNumber(property.Value);

                    if (found != null)
                        return found;
                }
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                {
                    var found = FindSerialNumber(item);

                    if (found != null)
                        return found;
                }
            }

            return null;
        }

        //Kiolvassa az új SPA felületű Bizhub számlálóit.
        static async Task<CounterInfo?> ReadSpaPrinter(string ip)
        {
            try
            {
                string baseUrl = $"http://{ip}";

                await client.GetAsync($"{baseUrl}/wcd/");
                await client.GetAsync($"{baseUrl}/wcd/index.html");
                await client.GetAsync($"{baseUrl}/wcd/spa_main.html");

                var baseUri = new Uri(baseUrl);


                // SPA indítás
                await client.GetAsync(
                    $"{baseUrl}/wcd/spa_main.html");


                // Inicializáló hívás
                var initResponse = await client.GetAsync(
                    $"{baseUrl}/wcd/api/AppReqGetCustomData/_A-00-00001?_=1");

                string initJson = await initResponse.Content.ReadAsStringAsync();

                string? serial = FindSerialNumber(JsonDocument.Parse(initJson).RootElement);


                await client.GetAsync(
                    $"{baseUrl}/wcd/trackinfo.xml?_=2");


                await client.GetAsync(
                    $"{baseUrl}/wcd/userinfo.xml?_=3");


                foreach (Cookie c in cookies.GetCookies(baseUri))
                {
                    Console.WriteLine(
                        $"{c.Name}={c.Value}");
                }



                string url =
                    $"{baseUrl}/wcd/api/AppReqGetCounterInfo/_Total";


                using var request = new HttpRequestMessage(
                    HttpMethod.Post,
                    url);


                request.Headers.Add(
                    "User-Agent",
                    "Mozilla/5.0");

                request.Headers.Add(
                    "X-Requested-With",
                    "XMLHttpRequest");


                request.Headers.Referrer =
                    new Uri($"{baseUrl}/wcd/spa_main.html");

                var page = await client.GetAsync(
                    $"{baseUrl}/wcd/spa_main.html");



                request.Content = new StringContent(
                    "",
                    Encoding.UTF8,
                    "application/json");



                using var response =
                    await client.SendAsync(request);


                string json =
                    await response.Content.ReadAsStringAsync();


                return ParseSpaCounter(json, serial);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        //A SPA API által visszaadott JSON feldolgozása.
        static CounterInfo ParseSpaCounter(string json, string serial)
        {
            using JsonDocument doc = JsonDocument.Parse(json);

            var root = doc.RootElement
                .GetProperty("MFP")
                .GetProperty("UserCounterInfo");


            int black = 0;
            int color = 0;


            // Nyomtatás
            var printList = root
                .GetProperty("PrintCounterList")
                .GetProperty("PrintCounter");


            foreach (var item in printList.EnumerateArray())
            {
                string type = item
                    .GetProperty("Type")
                    .GetString()!;

                int count = int.Parse(
                    item.GetProperty("Count").GetString()!);


                if (type == "BwTotal")
                    black += count;

                if (type == "FullColorTotal")
                    color += count;
            }


            // Másolás
            var copyList = root
                .GetProperty("CopyCounterList")
                .GetProperty("CopyCounter");


            foreach (var item in copyList.EnumerateArray())
            {
                string type = item
                    .GetProperty("Type")
                    .GetString()!;

                int count = int.Parse(
                    item.GetProperty("Count").GetString()!);


                if (type == "BwTotal")
                    black += count;

                if (type == "FullColorTotal")
                    color += count;
            }


            // Scan
            int scans = 0;

            var scanList = root
                .GetProperty("ScanFaxCounterList")
                .GetProperty("ScanFaxCounter");


            foreach (var item in scanList.EnumerateArray())
            {
                string type = item
                    .GetProperty("Type")
                    .GetString()!;


                if (type == "DocumentReadTotal")
                {
                    scans = int.Parse(
                        item.GetProperty("Count").GetString()!);
                }
            }


            return new CounterInfo
            {
                SerialNumber = serial,
                PrintBlack = black,
                PrintColor = color,
                Scans = scans
            };
        }

        //Egy adott XML elem számlálóértékének kiolvasása.
        static int GetValue(XDocument doc, string node, string type)
        {
            return doc.Descendants(node)
                .Where(x => x.Element("Type")?.Value == type)
                .Select(x => int.Parse(x.Element("Count")!.Value))
                .FirstOrDefault();
        }

        //Összeállítja az előző havi Excel fájl nevét.
        static string GetPreviousMonthFile()
        {
            DateTime previousMonth = DateTime.Now.AddMonths(-1);

            return Path.Combine(
                "Archiv",
                $"Számlálóállások - Colorspectrum - NSZFH_{previousMonth:yyyy_MM}.xlsx");
        }

        //Kiolvasás az excel cellákból.
        static int GetIntValue(IXLCell cell)
        {
            if (cell.IsEmpty())
                return 0;

            if (int.TryParse(cell.GetString(), out int value))
                return value;

            return 0;
        }

        //Beolvassa az előző havi Excel fájlt.
        static Dictionary<string, PreviousCounter> LoadPreviousCounters()
        {
            var previousCounters = new Dictionary<string, PreviousCounter>();

            string file = GetPreviousMonthFile();

            if (!File.Exists(file))
            {
                Console.WriteLine("Nincs előző havi fájl: " + file);
                return previousCounters;
            }

            using var workbook = new XLWorkbook(file);
            var ws = workbook.Worksheet(1);

            foreach (var row in ws.RowsUsed().Skip(2))
            {
                string serial = row.Cell(5).GetString().Trim();

                if (string.IsNullOrWhiteSpace(serial))
                    continue;

                previousCounters[serial] = new PreviousCounter
                {
                    Black = GetIntValue(row.Cell(9)),   // I - Jelenlegi FF
                    Color = GetIntValue(row.Cell(11)),  // K - Jelenlegi Sz
                    Scan = GetIntValue(row.Cell(13))    // M - Jelenlegi Scan
                };
            }

            Console.WriteLine($"Előző havi adatok betöltve: {previousCounters.Count} db");

            return previousCounters;
        }

        static void SaveResultsExcel(List<PrinterResult> printers, Dictionary<string, PreviousCounter> previousCounters)
        {
            using var workbook = new XLWorkbook("Archiv/Sablon.xlsx");
            var ws = workbook.Worksheet(1);

            // Gyors kikeresés sorozatszám alapján
            var resultsBySerial = printers
                .Where(p => !string.IsNullOrWhiteSpace(p.Serial))
                .ToDictionary(p => p.Serial.Trim(), p => p);

            foreach (var row in ws.RowsUsed().Skip(2))
            {
                string serial = row.Cell(5).GetString().Trim();

                if (string.IsNullOrWhiteSpace(serial))
                    continue;

                if (previousCounters.TryGetValue(serial, out var old))
                {
                    row.Cell(8).Value = old.Black;   // H - Előző FF
                    row.Cell(10).Value = old.Color;  // J - Előző Sz
                    row.Cell(12).Value = old.Scan;   // L - Előző Scan
                }

                if (resultsBySerial.TryGetValue(serial, out var current) && current.Counter != null)
                {
                    row.Cell(9).Value = current.Counter.PrintBlack;   // I - Jelenlegi FF
                    row.Cell(11).Value = current.Counter.PrintColor;  // K - Jelenlegi Sz
                    row.Cell(13).Value = current.Counter.Scans;       // M - Jelenlegi Scan
                    row.Cell(14).Value = "";
                }
                else
                {
                    row.Cell(9).Value = "-";
                    row.Cell(11).Value = "-";
                    row.Cell(13).Value = "-";
                    row.Cell(14).Value = "Jelenleg üzemen kívül";
                }
            }

            Directory.CreateDirectory("Archiv");

            string fileName = Path.Combine(
                "Archiv",
                $"Számlálóállások - Colorspectrum - NSZFH_{DateTime.Now:yyyy_MM}.xlsx");

            workbook.SaveAs(fileName);

            Console.WriteLine($"Mentve: {fileName}");
        }
    }
}