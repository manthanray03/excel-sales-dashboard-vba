// =====================================================================
//  SalesStaging  (Power Query / M)
//
//  "T" step of the ETL. Takes the raw POS rows that the VBA importer
//  loaded into tblRaw, cleans + standardizes them and enriches every
//  line with product and province details from the lookup tables.
//  Loads to tblStaging on the Staging sheet.
//
//  Fixable problems are fixed here (formats, casing, spacing, duplicate
//  exports). Anything that can't be fixed is left as null/blank so the
//  VBA validation step can flag it and route it to the Exceptions sheet.
// =====================================================================
let
    // ---------------- helper functions ----------------

    // tblRaw is text-formatted, but don't assume it
    ToText = (v as any) as text =>
        if v = null then ""
        else if Value.Is(v, type date) or Value.Is(v, type datetime) then
            Date.ToText(Date.From(v), "yyyy-MM-dd")
        else Text.Trim(Text.From(v, "en-US")),

    // "  john    smith " -> "john smith"
    CleanSpaces = (t as text) as text =>
        Text.Combine(List.Select(Text.Split(Text.Trim(t), " "), each _ <> ""), " "),

    // capitalise each word: "st. john's" -> "St. John's", "SMALL BUSINESS" -> "Small Business"
    ProperCase = (t as text) as text =>
        Text.Combine(
            List.Transform(
                Text.Split(Text.Lower(CleanSpaces(t)), " "),
                each if _ = "" then "" else Text.Upper(Text.Start(_, 1)) & Text.Middle(_, 1)),
            " "),

    // key used to match province spellings: "B.C." -> "BC", "Nova Scotia" -> "NOVASCOTIA"
    NormalizeKey = (t as text) as text =>
        Text.Upper(Text.Remove(t, {" ", ".", "-", "'"})),

    // most exports use YYYY-MM-DD, a few store terminals send MM/DD/YYYY
    ParseDate = (t as text) as nullable date =>
        let
            s = Text.Trim(t),
            isUS = Text.Contains(s, "/"),
            parts =
                if isUS then Text.Split(s, "/")
                else if Text.Contains(s, "-") then Text.Split(Text.Start(s, 10), "-")
                else {},
            result =
                if List.Count(parts) <> 3 then null
                else if isUS then
                    (try #date(Number.FromText(parts{2}, "en-US"), Number.FromText(parts{0}, "en-US"), Number.FromText(parts{1}, "en-US")) otherwise null)
                else
                    (try #date(Number.FromText(parts{0}, "en-US"), Number.FromText(parts{1}, "en-US"), Number.FromText(parts{2}, "en-US")) otherwise null)
        in
            result,

    // "$1,234.50" / " 3 " / "2.0" -> number, junk -> null
    ParseNumber = (t as text) as nullable number =>
        let s = Text.Remove(t, {"$", ",", " "})
        in if s = "" then null else (try Number.FromText(s, "en-US") otherwise null),

    // online store sends "10%", stores send 0.1 or leave it blank when there is no discount
    ParseDiscount = (t as text) as nullable number =>
        let s = Text.Remove(t, {" "})
        in
            if s = "" then 0
            else if Text.EndsWith(s, "%") then
                (try Number.FromText(Text.Start(s, Text.Length(s) - 1), "en-US") / 100 otherwise null)
            else
                (try Number.FromText(s, "en-US") otherwise null),

    StandardChannel = (t as text) as nullable text =>
        let k = Text.Lower(Text.Remove(t, {" ", "-", "_"}))
        in
            if List.Contains({"instore", "store", "pos"}, k) then "In-Store"
            else if List.Contains({"online", "web", "ecommerce"}, k) then "Online"
            else null,

    // ---------------- sources (all tables in this workbook) ----------------
    Source        = Excel.CurrentWorkbook(),
    RawTable      = Source{[Name = "tblRaw"]}[Content],
    ProductTable  = Source{[Name = "tblProducts"]}[Content],
    ProvinceTable = Source{[Name = "tblProvinces"]}[Content],
    AliasTable    = Source{[Name = "tblProvinceAlias"]}[Content],

    // ---------------- clean the raw rows ----------------
    RawColumns = {"Raw Row ID", "Source File", "Order ID", "Order Date", "Ship Date", "Customer ID",
                  "Customer Name", "Segment", "City", "Province", "Product ID", "Product Name",
                  "Quantity", "Unit Price", "Discount", "Sales Channel", "Sales Rep", "Payment Method"},
    SelectRaw = Table.SelectColumns(RawTable, RawColumns),
    AllText   = Table.TransformColumns(SelectRaw, List.Transform(RawColumns, (c) => {c, ToText, type text})),

    // blank lines at the end of some exports
    NonBlank = Table.SelectRows(AllText, each [Order ID] <> "" or [Product ID] <> ""),

    Renamed = Table.RenameColumns(NonBlank, {
        {"Raw Row ID", "RawRowID"}, {"Source File", "SourceFile"}, {"Order ID", "OrderID"},
        {"Order Date", "OrderDateText"}, {"Ship Date", "ShipDateText"}, {"Customer ID", "CustomerID"},
        {"Customer Name", "CustomerName"}, {"Province", "ProvinceText"}, {"Product ID", "ProductID"},
        {"Product Name", "ProductNameRaw"}, {"Quantity", "QuantityText"}, {"Unit Price", "UnitPriceText"},
        {"Discount", "DiscountText"}, {"Sales Channel", "ChannelText"}, {"Sales Rep", "SalesRep"},
        {"Payment Method", "PaymentMethod"}}),

    Standardized = Table.TransformColumns(Renamed, {
        {"RawRowID", each (try Number.FromText(_, "en-US") otherwise null), Int64.Type},
        {"OrderID", Text.Upper, type text},
        {"CustomerID", Text.Upper, type text},
        {"CustomerName", ProperCase, type text},
        {"Segment", ProperCase, type text},
        {"City", ProperCase, type text},
        {"ProductID", Text.Upper, type text},
        {"SalesRep", CleanSpaces, type text},
        {"PaymentMethod", CleanSpaces, type text}}),

    AddOrderDate = Table.AddColumn(Standardized, "OrderDate", each ParseDate([OrderDateText]), type nullable date),
    AddShipDate  = Table.AddColumn(AddOrderDate, "ShipDate", each ParseDate([ShipDateText]), type nullable date),
    AddQuantity  = Table.AddColumn(AddShipDate, "Quantity", each ParseNumber([QuantityText]), type nullable number),
    AddUnitPrice = Table.AddColumn(AddQuantity, "UnitPrice", each ParseNumber([UnitPriceText]), type nullable number),
    AddDiscount  = Table.AddColumn(AddUnitPrice, "Discount", each ParseDiscount([DiscountText]), type nullable number),
    AddChannel   = Table.AddColumn(AddDiscount, "Channel", each StandardChannel([ChannelText]), type nullable text),
    AddProvKey   = Table.AddColumn(AddChannel, "ProvinceKey", each NormalizeKey([ProvinceText]), type text),

    // the same export sometimes gets loaded twice -> drop exact repeats, keep the first one
    Sorted    = Table.Buffer(Table.Sort(AddProvKey, {{"RawRowID", Order.Ascending}})),
    DedupKeys = {"OrderID", "OrderDate", "ShipDate", "CustomerID", "CustomerName", "Segment", "City",
                 "ProvinceKey", "ProductID", "Quantity", "UnitPrice", "Discount", "Channel",
                 "SalesRep", "PaymentMethod"},
    Deduped   = Table.Distinct(Sorted, DedupKeys),

    // ---------------- province lookup ----------------
    Aliases = Table.Distinct(
        Table.TransformColumns(
            Table.SelectColumns(AliasTable, {"Alias", "ProvinceCode"}),
            {{"Alias", each NormalizeKey(ToText(_)), type text},
             {"ProvinceCode", each Text.Upper(ToText(_)), type text}}),
        {"Alias"}),
    MergeAlias  = Table.NestedJoin(Deduped, {"ProvinceKey"}, Aliases, {"Alias"}, "AliasMatch", JoinKind.LeftOuter),
    ExpandAlias = Table.ExpandTableColumn(MergeAlias, "AliasMatch", {"ProvinceCode"}, {"ProvinceCode"}),

    Provinces = Table.TransformColumns(
        Table.SelectColumns(ProvinceTable, {"ProvinceCode", "Province", "Region"}),
        {{"ProvinceCode", each Text.Upper(ToText(_)), type text}}),
    MergeProv  = Table.NestedJoin(ExpandAlias, {"ProvinceCode"}, Provinces, {"ProvinceCode"}, "ProvMatch", JoinKind.LeftOuter),
    ExpandProv = Table.ExpandTableColumn(MergeProv, "ProvMatch", {"Province", "Region"}, {"Province", "Region"}),

    // ---------------- product lookup ----------------
    Products = Table.TransformColumns(
        Table.SelectColumns(ProductTable, {"ProductID", "ProductName", "Category", "SubCategory", "ListPrice", "UnitCost"}),
        {{"ProductID", each Text.Upper(ToText(_)), type text}}),
    MergeProd  = Table.NestedJoin(ExpandProv, {"ProductID"}, Products, {"ProductID"}, "ProdMatch", JoinKind.LeftOuter),
    ExpandProd = Table.ExpandTableColumn(MergeProd, "ProdMatch",
        {"ProductName", "Category", "SubCategory", "ListPrice", "UnitCost"},
        {"ProductName", "Category", "SubCategory", "ListPrice", "UnitCost"}),

    // use the product master name when we know the product, otherwise keep what the till printed
    AddName = Table.AddColumn(ExpandProd, "ProductNameClean",
        each if [ProductName] <> null then [ProductName] else [ProductNameRaw], type text),

    // ---------------- final shape ----------------
    Output = Table.SelectColumns(AddName, {
        "RawRowID", "SourceFile", "OrderID", "OrderDate", "ShipDate", "CustomerID", "CustomerName",
        "Segment", "City", "ProvinceCode", "Province", "Region", "ProductID", "ProductNameClean",
        "Category", "SubCategory", "Quantity", "UnitPrice", "ListPrice", "UnitCost", "Discount",
        "Channel", "SalesRep", "PaymentMethod"}),
    RenamedOut = Table.RenameColumns(Output, {{"ProductNameClean", "ProductName"}}),
    Typed = Table.TransformColumnTypes(RenamedOut, {
        {"RawRowID", Int64.Type}, {"OrderDate", type date}, {"ShipDate", type date},
        {"ProvinceCode", type text}, {"Province", type text}, {"Region", type text},
        {"Category", type text}, {"SubCategory", type text},
        {"Quantity", type number}, {"UnitPrice", type number}, {"ListPrice", type number},
        {"UnitCost", type number}, {"Discount", type number}}, "en-US"),
    Result = Table.Sort(Typed, {{"RawRowID", Order.Ascending}})
in
    Result
