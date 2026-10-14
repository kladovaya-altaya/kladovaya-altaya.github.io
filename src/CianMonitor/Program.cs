using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Microsoft.Playwright;

const string listingSearch = "https://www.cian.ru/cat.php?deal_type=sale&engine_version=2&id_user=118101459&offer_type=flat";
var root = Directory.GetCurrentDirectory();
var dataPath = Path.Combine(root, "data", "state.json");
var sitePath = Path.Combine(root, "cian-monitor");
var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower, PropertyNameCaseInsensitive = true, WriteIndented = true };
var state = JsonSerializer.Deserialize<State>(File.ReadAllText(dataPath), options) ?? new State();
var checkedAt = DateTimeOffset.UtcNow.ToString("O");
ScanStatus status;
try
{
    var scan = await ScanAsync(listingSearch);
    var previousActive = state.Listings.Values.Count(x => x.Status.StartsWith("В профиле", StringComparison.Ordinal));
    if (previousActive >= 2 && scan.SeenIds.Count < previousActive / 2)
        throw new InvalidOperationException($"Список подозрительно сократился: было {previousActive}, найдено {scan.SeenIds.Count}. Статусы сохранены до следующей проверки");
    foreach (var (id, item) in scan.Items)
    {
        state.Listings.TryGetValue(id, out var old);
        item.Id = id;
        item.FirstSeen = old?.FirstSeen ?? checkedAt;
        item.LastSeen = checkedAt;
        item.Status = "В профиле";
        item.PriceHistory = old?.PriceHistory ?? [];
        if (item.PriceRub is null) item.PriceRub = old?.PriceRub;
        if (item.PriceRub is not null && (item.PriceHistory.Count == 0 || item.PriceHistory[^1].PriceRub != item.PriceRub))
            item.PriceHistory.Add(new PricePoint { At = checkedAt, PriceRub = item.PriceRub.Value });
        state.Listings[id] = item;
    }
    foreach (var id in scan.SeenIds.Except(scan.Items.Keys))
    {
        if (state.Listings.TryGetValue(id, out var item))
        {
            item.LastSeen = checkedAt;
            item.Status = "В профиле, детали недоступны";
        }
    }
    foreach (var (id, item) in state.Listings)
    {
        if (!scan.SeenIds.Contains(id)) item.Status = "Нет в профиле";
    }
    state.LastSuccess = checkedAt;
    Directory.CreateDirectory(Path.GetDirectoryName(dataPath)!);
    File.WriteAllText(dataPath, JsonSerializer.Serialize(state, options));
    status = new ScanStatus { CheckedAt = checkedAt, Ok = true, Found = scan.SeenIds.Count, DetailErrors = scan.Errors };
}
catch (Exception ex)
{
    status = new ScanStatus { CheckedAt = checkedAt, Ok = false, Error = ex.Message, LastSuccess = state.LastSuccess };
    Console.Error.WriteLine(ex.Message);
}

Directory.CreateDirectory(sitePath);
foreach (var item in state.Listings.Values)
    item.DaysObserved = Math.Max(1, (DateTimeOffset.UtcNow.Date - DateTimeOffset.Parse(item.FirstSeen).UtcDateTime.Date).Days + 1);
File.WriteAllText(Path.Combine(sitePath, "listings.json"), JsonSerializer.Serialize(state.Listings.Values.OrderByDescending(x => x.FirstSeen), options));
File.WriteAllText(Path.Combine(sitePath, "status.json"), JsonSerializer.Serialize(status, options));
WriteExcel(Path.Combine(sitePath, "cian-monitor.xlsx"), state.Listings.Values);
return status.Ok ? 0 : 1;

static async Task<ScanResult> ScanAsync(string searchUrl)
{
    using var playwright = await Playwright.CreateAsync();
    await using var browser = await playwright.Chromium.LaunchAsync(new() { Headless = true });
    var page = await browser.NewPageAsync(new() { Locale = "ru-RU", ViewportSize = new() { Width = 1365, Height = 900 } });
    await ReadAsync(page, searchUrl);
    var links = new Dictionary<string, string>();
    for (var attempt = 0; attempt < 12; attempt++)
    {
        var hrefs = await page.Locator("a[href]").EvaluateAllAsync<string[]>("els => els.map(a => a.href)");
        foreach (var href in hrefs)
        {
            if (!Uri.TryCreate(href, UriKind.Absolute, out var uri) || !uri.Host.EndsWith("cian.ru")) continue;
            var m = Regex.Match(uri.AbsolutePath, @"/(?:sale|rent)/(?:flat|suburban|commercial)/(\d+)/?", RegexOptions.IgnoreCase);
            if (m.Success) links[m.Groups[1].Value] = uri.GetLeftPart(UriPartial.Path);
        }
        await page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");
        await page.WaitForTimeoutAsync(1500);
        var more = page.GetByText(new Regex("Показать ещё|Показать еще|Загрузить ещё", RegexOptions.IgnoreCase));
        if (await more.CountAsync() == 0 || !await more.First.IsVisibleAsync()) break;
        await more.First.ClickAsync(new() { Timeout = 3000 });
        await page.WaitForTimeoutAsync(900);
    }
    if (links.Count == 0)
    {
        var responseTitle = await page.TitleAsync();
        var responseBody = await page.Locator("body").InnerTextAsync();
        var excerpt = Regex.Replace(responseBody, @"\s+", " ").Trim();
        if (excerpt.Length > 180) excerpt = excerpt[..180];
        throw new InvalidOperationException($"В результатах поиска не найдены ссылки на объявления. Страница: {page.Url}; заголовок: {responseTitle}; начало текста: {excerpt}");
    }
    var searchText = await page.Locator("body").InnerTextAsync();
    var expectedMatch = Regex.Match(searchText, @"Найдено\s+(\d[\d\s]*)\s+объявлен", RegexOptions.IgnoreCase);
    if (expectedMatch.Success && int.TryParse(Regex.Replace(expectedMatch.Groups[1].Value, @"\s", ""), out var expected) && links.Count < expected)
        throw new InvalidOperationException($"Список загружен не полностью: Циан сообщает {expected} объявлений, найдено ссылок {links.Count}. Статусы сохранены до следующей проверки");
    var result = new ScanResult { SeenIds = links.Keys.ToHashSet() };
    foreach (var (id, url) in links)
    {
        try
        {
            var body = await ReadAsync(page, url);
            var title = await page.TitleAsync();
            var description = await page.Locator("meta[name='description']").GetAttributeAsync("content") ?? "";
            result.Items[id] = Extract(url, title, body, description);
        }
        catch (Exception ex) { result.Errors.Add($"{id}: {ex.Message}"); }
    }
    return result;
}

static async Task<string> ReadAsync(IPage page, string url)
{
    await page.GotoAsync(url, new() { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 45000 });
    await page.WaitForTimeoutAsync(2500);
    var title = await page.TitleAsync();
    if (Regex.IsMatch(page.Url + " " + title, "showcaptcha|вы не робот|captcha", RegexOptions.IgnoreCase))
        throw new InvalidOperationException("Циан запросил капчу");
    var body = await page.Locator("body").InnerTextAsync(new() { Timeout = 10000 });
    if (string.IsNullOrWhiteSpace(body)) throw new InvalidOperationException("Пустая страница Циан");
    return body;
}

static Listing Extract(string url, string title, string body, string description)
{
    var area = Find(@"Общая площадь\s*([\d,.]+)\s*м²", body);
    if (area == "") area = Find(@"([\d,.]+)\s*м²", title);
    var price = Find(@"([\d\s\u00a0]+)\s*₽", body);
    var lines = body.Split('\n', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    string metro = "";
    int? metroMinutes = null;
    for (var i = 1; i < lines.Length; i++)
    {
        var m = Regex.Match(lines[i], @"^(\d+)\s*мин\.?(?:\s|$)", RegexOptions.IgnoreCase);
        if (!m.Success) continue;
        if (lines[i - 1].Length > 45 || Regex.IsMatch(lines[i - 1], @"\d+\s*км|МКАД", RegexOptions.IgnoreCase)) continue;
        metro = lines[i - 1];
        metroMinutes = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
        break;
    }
    return new Listing
    {
        Url = url, Title = title, Description = description.Length > 2000 ? description[..2000] : description,
        Kind = Find(@"(студия|\d+-комн\.\s*квартира|апартаменты)", title),
        Address = Find(@"((?:Москва|Санкт-Петербург|Московская область)[^\n]{0,150})", body),
        AreaM2 = ParseDouble(area), Floor = Find(@"Этаж\s*(\d+\s*из\s*\d+)", body),
        Metro = metro, MetroMinutes = metroMinutes, PriceRub = ParseLong(price),
        CianUpdated = Find(@"Обновлено:\s*([^\n]+)", body)
    };
}

static string Find(string pattern, string text)
{
    var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
    return match.Success ? match.Groups[1].Value.Trim() : "";
}
static double? ParseDouble(string value) => double.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
static long? ParseLong(string value) => long.TryParse(Regex.Replace(value, @"\s", ""), out var parsed) ? parsed : null;

static void WriteExcel(string path, IEnumerable<Listing> listings)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.Worksheets.Add("Объявления");
    string[] headers = ["ID", "Объявление", "Тип", "Адрес", "Площадь, м²", "Этаж", "Метро", "До метро, мин", "Первая цена, ₽", "Цена сейчас, ₽", "Статус", "Впервые замечено", "Последний раз замечено", "Дней наблюдения", "Обновлено на Циан", "Ссылка"];
    for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
    var row = 2;
    foreach (var x in listings.OrderByDescending(x => x.FirstSeen))
    {
        sheet.Cell(row, 1).Value = x.Id;
        sheet.Cell(row, 2).Value = x.Title;
        sheet.Cell(row, 3).Value = x.Kind;
        sheet.Cell(row, 4).Value = x.Address;
        if (x.AreaM2 is double area) sheet.Cell(row, 5).Value = area;
        sheet.Cell(row, 6).Value = x.Floor;
        sheet.Cell(row, 7).Value = x.Metro;
        if (x.MetroMinutes is int minutes) sheet.Cell(row, 8).Value = minutes;
        if (x.PriceHistory.FirstOrDefault() is { } firstPrice) sheet.Cell(row, 9).Value = firstPrice.PriceRub;
        if (x.PriceRub is long price) sheet.Cell(row, 10).Value = price;
        sheet.Cell(row, 11).Value = x.Status;
        sheet.Cell(row, 12).Value = x.FirstSeen;
        sheet.Cell(row, 13).Value = x.LastSeen;
        sheet.Cell(row, 14).Value = x.DaysObserved;
        sheet.Cell(row, 15).Value = x.CianUpdated;
        sheet.Cell(row, 16).Value = x.Url;
        row++;
    }
    var range = sheet.Range(1, 1, Math.Max(1, row - 1), headers.Length);
    range.SetAutoFilter();
    sheet.SheetView.FreezeRows(1);
    sheet.Row(1).Style.Font.Bold = true;
    sheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF1FA");
    sheet.Columns(1, headers.Length).AdjustToContents();
    foreach (var c in new[] { 2, 4, 7, 16 }) if (sheet.Column(c).Width > 55) sheet.Column(c).Width = 55;
    workbook.SaveAs(path);
}

public sealed class State
{
    public Dictionary<string, Listing> Listings { get; set; } = [];
    public string? LastSuccess { get; set; }
}
public sealed class Listing
{
    public string Id { get; set; } = "";
    public string Url { get; set; } = "";
    public string Title { get; set; } = "";
    public string Kind { get; set; } = "";
    public string Address { get; set; } = "";
    public double? AreaM2 { get; set; }
    public string Floor { get; set; } = "";
    public string Metro { get; set; } = "";
    public int? MetroMinutes { get; set; }
    public long? PriceRub { get; set; }
    public string CianUpdated { get; set; } = "";
    public string Description { get; set; } = "";
    public string FirstSeen { get; set; } = "";
    public string LastSeen { get; set; } = "";
    public string Status { get; set; } = "";
    public int DaysObserved { get; set; }
    public List<PricePoint> PriceHistory { get; set; } = [];
}
public sealed class PricePoint
{
    public string At { get; set; } = "";
    public long PriceRub { get; set; }
}
public sealed class ScanStatus
{
    public string CheckedAt { get; set; } = "";
    public bool Ok { get; set; }
    public int Found { get; set; }
    public List<string> DetailErrors { get; set; } = [];
    public string? Error { get; set; }
    public string? LastSuccess { get; set; }
}
public sealed class ScanResult
{
    public HashSet<string> SeenIds { get; set; } = [];
    public Dictionary<string, Listing> Items { get; set; } = [];
    public List<string> Errors { get; set; } = [];
}
