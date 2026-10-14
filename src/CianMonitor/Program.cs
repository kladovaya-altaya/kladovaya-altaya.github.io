using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

const string agentId = "118101459";
var root = Directory.GetCurrentDirectory();
var dataPath = Path.Combine(root, "data", "state.json");
var sitePath = Path.Combine(root, "cian-monitor");
var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower, PropertyNameCaseInsensitive = true, WriteIndented = true };
var state = JsonSerializer.Deserialize<State>(File.ReadAllText(dataPath), options) ?? new State();
var checkedAt = DateTimeOffset.UtcNow.ToString("O");
ScanStatus status;
try
{
    var scan = await ScanAsync(agentId);
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
        // A confirmed first publication date must not be replaced by a later refresh.
        if (!string.IsNullOrWhiteSpace(old?.PublishedAt)) item.PublishedAt = old.PublishedAt;
        if (string.IsNullOrWhiteSpace(item.PublishedAt)) item.PublishedAt = await TryReadCianPublicationAsync(id);
        if (string.IsNullOrWhiteSpace(item.MetroTransport)) item.MetroTransport = old?.MetroTransport ?? "";
        item.KitchenAreaM2 ??= old?.KitchenAreaM2;
        if (string.IsNullOrWhiteSpace(item.BalconyLoggia)) item.BalconyLoggia = old?.BalconyLoggia ?? "";
        item.HouseFloors ??= old?.HouseFloors;
        item.BuiltYear ??= old?.BuiltYear;
        if (string.IsNullOrWhiteSpace(item.BuildingType)) item.BuildingType = old?.BuildingType ?? "";
        if (string.IsNullOrWhiteSpace(item.GasSupply)) item.GasSupply = old?.GasSupply ?? "";
        item.CeilingHeightM ??= old?.CeilingHeightM;
        if (string.IsNullOrWhiteSpace(item.CianUpdated)) item.CianUpdated = old?.CianUpdated ?? "";
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
    status = new ScanStatus { CheckedAt = checkedAt, Ok = true, Found = scan.SeenIds.Count, DetailErrors = scan.Errors };
}
catch (Exception ex)
{
    status = new ScanStatus { CheckedAt = checkedAt, Ok = false, Error = ex.Message, LastSuccess = state.LastSuccess };
    Console.Error.WriteLine(ex.Message);
}

Directory.CreateDirectory(sitePath);
foreach (var item in state.Listings.Values)
{
    var moscow = TimeSpan.FromHours(3);
    item.DaysObserved = string.IsNullOrWhiteSpace(item.PublishedAt) ? 0 :
        Math.Max(1, (DateTimeOffset.UtcNow.ToOffset(moscow).Date - DateTimeOffset.Parse(item.PublishedAt).ToOffset(moscow).Date).Days + 1);
}
if (status.Ok) File.WriteAllText(dataPath, JsonSerializer.Serialize(state, options));
File.WriteAllText(Path.Combine(sitePath, "listings.json"), JsonSerializer.Serialize(state.Listings.Values.OrderByDescending(x => x.FirstSeen), options));
File.WriteAllText(Path.Combine(sitePath, "status.json"), JsonSerializer.Serialize(status, options));
WriteExcel(Path.Combine(sitePath, "cian-monitor.xlsx"), state.Listings.Values);
return status.Ok ? 0 : 1;

static async Task<ScanResult> ScanAsync(string agentId)
{
    var key = Environment.GetEnvironmentVariable("REALTYAPI_KEY");
    if (string.IsNullOrWhiteSpace(key)) throw new InvalidOperationException("Не задан секрет REALTYAPI_KEY");
    using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
    client.DefaultRequestHeaders.Add("x-realtyapi-key", key);
    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    var result = new ScanResult();
    int? total = null;
    for (var page = 1; page <= 100; page++)
    {
        var url = $"https://cian.realtyapi.io/agent/offers?id={agentId}&page={page}&resultCount=28";
        using var response = await client.GetAsync(url);
        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException($"RealtyAPI вернул HTTP {(int)response.StatusCode}; архив не изменён");
        using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var root = doc.RootElement;
        if (GetString(root, "message") != "Success" || !root.TryGetProperty("searchResults", out var offers) || offers.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("RealtyAPI вернул неожиданный ответ; архив не изменён");
        total ??= GetInt(root, "total") ?? throw new InvalidOperationException("RealtyAPI не указал общее число объявлений");
        foreach (var offer in offers.EnumerateArray())
        {
            var id = GetLong(offer, "id")?.ToString(CultureInfo.InvariantCulture);
            var owner = offer.TryGetProperty("agent", out var agent) ? GetString(agent, "id") : null;
            if (id is null || owner != agentId || GetString(offer, "dealType") != "sale" || GetString(offer, "offerType") != "flat")
                throw new InvalidOperationException("В ответе RealtyAPI обнаружено объявление другого агента или типа; архив не изменён");
            if (!result.SeenIds.Add(id)) throw new InvalidOperationException($"RealtyAPI повторил объявление {id}; архив не изменён");
            result.Items[id] = Extract(offer, id);
        }
        if (root.TryGetProperty("nextPage", out var next) && next.ValueKind == JsonValueKind.False) break;
        if (page == 100) throw new InvalidOperationException("Превышен лимит страниц RealtyAPI; архив не изменён");
    }
    if (result.SeenIds.Count != total)
        throw new InvalidOperationException($"RealtyAPI сообщил {total} объявлений, получено {result.SeenIds.Count}; архив не изменён");
    return result;
}

static Listing Extract(JsonElement offer, string id)
{
    var building = offer.TryGetProperty("building", out var b) ? b : default;
    var price = offer.TryGetProperty("price", out var p) ? GetLong(p, "value") : null;
    if (price is null || price <= 0) throw new InvalidOperationException($"У объявления {id} нет действительной цены");
    var rooms = GetInt(offer, "rooms");
    var kind = rooms is > 0 ? $"{rooms}-комн. квартира" : "Квартира";
    if (GetString(offer, "flatType") == "studio") kind = "Студия";
    var address = GetString(offer, "address") ?? "";
    var metro = "";
    int? metroMinutes = null;
    var metroTransport = "";
    if (offer.TryGetProperty("metro", out var stations) && stations.ValueKind == JsonValueKind.Array)
    {
        var first = stations.EnumerateArray().FirstOrDefault();
        if (first.ValueKind == JsonValueKind.Object)
        {
            metro = GetString(first, "name") ?? "";
            metroMinutes = GetInt(first, "timeMin");
            metroTransport = GetString(first, "transport") switch { "walk" => "пешком", "transport" => "на транспорте", _ => "" };
        }
    }
    return new Listing
    {
        Url = $"https://www.cian.ru/sale/flat/{id}/",
        Title = $"{kind}, {address}", Kind = kind, Address = address,
        Description = Truncate(GetString(offer, "description") ?? "", 2000),
        AreaM2 = GetDouble(offer, "totalArea"),
        KitchenAreaM2 = GetDouble(offer, "kitchenArea"),
        Floor = GetInt(offer, "floor") is int floor && GetInt(offer, "floorsCount") is int floors ? $"{floor} из {floors}" : "",
        HouseFloors = GetInt(building, "floorsCount") ?? GetInt(offer, "floorsCount"),
        BuiltYear = GetInt(building, "buildYear"),
        BuildingType = GetString(building, "material") switch { "panel" => "Панельный", "brick" => "Кирпичный", "monolith" => "Монолитный", "block" => "Блочный", var material => material ?? "" },
        Metro = metro, MetroMinutes = metroMinutes, MetroTransport = metroTransport, PriceRub = price,
        PublishedAt = GetLong(offer, "publicationDate") is long published ? DateTimeOffset.FromUnixTimeSeconds(published).ToOffset(TimeSpan.FromHours(3)).ToString("O") : "",
        CianUpdated = GetLong(offer, "addedTimestamp") is long updated ? DateTimeOffset.FromUnixTimeSeconds(updated).ToOffset(TimeSpan.FromHours(3)).ToString("O") : ""
    };
}

static async Task<string> TryReadCianPublicationAsync(string id)
{
    try
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; CianMonitor/1.0)");
        var html = await client.GetStringAsync($"https://spb.cian.ru/sale/flat/{id}/");
        var match = Regex.Match(html, "\"publicationDate\":(?<date>\\d+),\"creationDate\":\"");
        if (match.Success && long.TryParse(match.Groups["date"].Value, out var timestamp))
            return DateTimeOffset.FromUnixTimeSeconds(timestamp).ToOffset(TimeSpan.FromHours(3)).ToString("O");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Исходная дата объявления {id} недоступна: {ex.GetType().Name}");
    }
    return "";
}

static string Truncate(string value, int limit) => value.Length <= limit ? value : value[..limit];
static string? GetString(JsonElement item, string key) => item.ValueKind == JsonValueKind.Object && item.TryGetProperty(key, out var value) && value.ValueKind != JsonValueKind.Null ? value.ToString() : null;
static long? GetLong(JsonElement item, string key) => item.ValueKind == JsonValueKind.Object && item.TryGetProperty(key, out var value) && value.TryGetInt64(out var number) ? number : null;
static int? GetInt(JsonElement item, string key) => item.ValueKind == JsonValueKind.Object && item.TryGetProperty(key, out var value) && value.TryGetInt32(out var number) ? number : null;
static double? GetDouble(JsonElement item, string key) => item.ValueKind == JsonValueKind.Object && item.TryGetProperty(key, out var value) && value.TryGetDouble(out var number) ? number : null;

static void WriteExcel(string path, IEnumerable<Listing> listings)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.Worksheets.Add("Объявления");
    string[] headers = ["ID", "Объявление", "Тип", "Адрес", "Площадь, м²", "Кухня, м²", "Балкон/лоджия", "Этаж", "Этажей в доме", "Год постройки", "Тип дома", "Газоснабжение", "Потолки, м", "Метро", "До метро", "Первая цена, ₽", "Цена сейчас, ₽", "Статус", "Опубликовано на Циане", "Впервые замечено монитором", "Последний раз замечено", "Дней на Циане", "Дата addedTimestamp из API (не первая публикация)", "Ссылка"];
    for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
    var row = 2;
    foreach (var x in listings.OrderByDescending(x => x.FirstSeen))
    {
        sheet.Cell(row, 1).Value = x.Id;
        sheet.Cell(row, 2).Value = x.Title;
        sheet.Cell(row, 3).Value = x.Kind;
        sheet.Cell(row, 4).Value = x.Address;
        if (x.AreaM2 is double area) sheet.Cell(row, 5).Value = area;
        if (x.KitchenAreaM2 is double kitchen) sheet.Cell(row, 6).Value = kitchen;
        sheet.Cell(row, 7).Value = x.BalconyLoggia;
        sheet.Cell(row, 8).Value = x.Floor;
        if (x.HouseFloors is int houseFloors) sheet.Cell(row, 9).Value = houseFloors;
        if (x.BuiltYear is int builtYear) sheet.Cell(row, 10).Value = builtYear;
        sheet.Cell(row, 11).Value = x.BuildingType;
        sheet.Cell(row, 12).Value = x.GasSupply;
        if (x.CeilingHeightM is double ceiling) sheet.Cell(row, 13).Value = ceiling;
        sheet.Cell(row, 14).Value = x.Metro;
        if (x.MetroMinutes is int minutes) sheet.Cell(row, 15).Value = $"{minutes} мин {x.MetroTransport}".Trim();
        if (x.PriceHistory.FirstOrDefault() is { } firstPrice) sheet.Cell(row, 16).Value = firstPrice.PriceRub;
        if (x.PriceRub is long price) sheet.Cell(row, 17).Value = price;
        sheet.Cell(row, 18).Value = x.Status;
        sheet.Cell(row, 19).Value = x.PublishedAt;
        sheet.Cell(row, 20).Value = x.FirstSeen;
        sheet.Cell(row, 21).Value = x.LastSeen;
        if (x.DaysObserved > 0) sheet.Cell(row, 22).Value = x.DaysObserved;
        sheet.Cell(row, 23).Value = x.CianUpdated;
        sheet.Cell(row, 24).Value = x.Url;
        row++;
    }
    var range = sheet.Range(1, 1, Math.Max(1, row - 1), headers.Length);
    range.SetAutoFilter();
    sheet.SheetView.FreezeRows(1);
    sheet.Row(1).Style.Font.Bold = true;
    sheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF1FA");
    sheet.Columns(1, headers.Length).AdjustToContents();
    foreach (var c in new[] { 2, 4, 14, 24 }) if (sheet.Column(c).Width > 55) sheet.Column(c).Width = 55;
    var history = workbook.Worksheets.Add("История цен");
    string[] historyHeaders = ["ID", "Объявление", "Дата наблюдения", "Цена, ₽", "Ссылка"];
    for (var c = 0; c < historyHeaders.Length; c++) history.Cell(1, c + 1).Value = historyHeaders[c];
    var historyRow = 2;
    foreach (var x in listings)
        foreach (var point in x.PriceHistory.OrderBy(p => p.At))
        {
            history.Cell(historyRow, 1).Value = x.Id;
            history.Cell(historyRow, 2).Value = x.Title;
            history.Cell(historyRow, 3).Value = point.At;
            history.Cell(historyRow, 4).Value = point.PriceRub;
            history.Cell(historyRow, 5).Value = x.Url;
            historyRow++;
        }
    history.Range(1, 1, Math.Max(1, historyRow - 1), historyHeaders.Length).SetAutoFilter();
    history.SheetView.FreezeRows(1);
    history.Row(1).Style.Font.Bold = true;
    history.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF1FA");
    history.Columns(1, historyHeaders.Length).AdjustToContents();
    if (history.Column(2).Width > 55) history.Column(2).Width = 55;
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
    public double? KitchenAreaM2 { get; set; }
    public string BalconyLoggia { get; set; } = "";
    public string Floor { get; set; } = "";
    public int? HouseFloors { get; set; }
    public int? BuiltYear { get; set; }
    public string BuildingType { get; set; } = "";
    public string GasSupply { get; set; } = "";
    public double? CeilingHeightM { get; set; }
    public string Metro { get; set; } = "";
    public int? MetroMinutes { get; set; }
    public string MetroTransport { get; set; } = "";
    public long? PriceRub { get; set; }
    public string CianUpdated { get; set; } = "";
    public string Description { get; set; } = "";
    public string FirstSeen { get; set; } = "";
    public string PublishedAt { get; set; } = "";
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
