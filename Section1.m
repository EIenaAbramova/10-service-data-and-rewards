section Section1;

shared SourceFolder = Text.From(Excel.CurrentWorkbook(){[Name="SourceFolder"]}[Content]{0}[Column1]);

shared CleanText = (x as any) as text =>
    if x = null then "" else Text.Combine(List.Select(Text.SplitAny(Text.Replace(Text.Clean(Text.From(x)), Character.FromNumber(160), " "), " "), each _ <> ""), " ");

shared ReadFile = (content as binary, fileName as text) as table =>
let
    Book = Excel.Workbook(content, null, true),
    Sheet = Book{[Item="Services", Kind="Sheet"]}[Data],
    Headers = Table.TransformColumnNames(Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]), each Text.Lower(CleanText(_))),
    Indexed = Table.AddIndexColumn(Headers, "Строка", 2, 1, Int64.Type),
    Selected = Table.SelectColumns(Indexed, {"Строка", "id пациента", "возраст", "дата услуги", "клиника услуги", "компенсация, руб.", "представитель"}),
    NonEmpty = Table.SelectRows(Selected, each not List.IsEmpty(List.RemoveNulls(Record.ToList(Record.RemoveFields(_, {"Строка"}))))),
    WithFile = Table.AddColumn(NonEmpty, "Файл", each fileName, type text)
in WithFile;

shared Услуги =
let
    Files = Table.SelectRows(Folder.Contents(SourceFolder), each Text.StartsWith([Name], "Services_") and Text.Lower([Extension]) = ".xlsx" and [Attributes]?[Hidden]? <> true),
    Combined = if Table.RowCount(Files)=0 then error "Нет файлов Services_*.xlsx в выбранной папке" else Table.Combine(List.Transform(Table.ToRecords(Files), each ReadFile([Content], [Name]))),
    Typed = Table.TransformColumns(Combined, {
        {"id пациента", each if CleanText(_)="" then error "Нет id пациента" else CleanText(_), type text},
        {"возраст", each if _=null then error "Нет возраста" else Number.From(_, "ru-RU"), type number},
        {"дата услуги", each if _=null then error "Нет даты" else Date.From(_, "ru-RU"), type date},
        {"клиника услуги", each if CleanText(_)="" then error "Нет клиники" else CleanText(_), type text},
        {"компенсация, руб.", each if _=null then error "Нет компенсации" else Currency.From(_, "ru-RU"), Currency.Type},
        {"представитель", each CleanText(_), type text}
    }),
    Checked = if Table.RowCount(Table.SelectRows(Typed, each [возраст]<0 or [возраст]>120))>0 then error "Некорректный возраст" else Typed,
    Month = Table.AddColumn(Checked, "Месяц", each Date.ToText([дата услуги], "yyyy-MM"), type text),
    Age = Table.AddColumn(Month, "Возрастная категория", each if [возраст]>=18 then "взрослые" else "дети", type text),
    Channel = Table.AddColumn(Age, "Канал привлечения", each if [представитель]<>"" then "от Агентов" else "Реклама", type text),
    Result = Table.ReorderColumns(Channel, {"Файл","Строка","id пациента","возраст","дата услуги","Месяц","клиника услуги","компенсация, руб.","представитель","Возрастная категория","Канал привлечения"})
in Result;

shared Уникальные =
let
    Keys = Table.Distinct(Table.SelectColumns(Услуги, {"Месяц","клиника услуги","Канал привлечения","id пациента"})),
    Unit = Table.AddColumn(Keys, "Единица", each 1, Int64.Type)
in Table.Sort(Unit, {{"Месяц",Order.Ascending},{"клиника услуги",Order.Ascending},{"Канал привлечения",Order.Ascending},{"id пациента",Order.Ascending}});

shared ПациентМесяц =
let
    Grouped = Table.Group(Услуги, {"Месяц","id пациента"}, {
        {"представитель", each let r=List.Distinct(List.Select([представитель], each _<>"")) in if List.Count(r)>1 then error "Несколько представителей за месяц" else if List.IsEmpty(r) then "" else r{0}, type text},
        {"Возрастная категория", each let c=List.Distinct([Возрастная категория]) in if List.Count(c)<>1 then error "Возрастная категория изменилась за месяц" else c{0}, type text},
        {"Оплата за месяц, руб.", each List.Sum([#"компенсация, руб."]), Currency.Type},
        {"Строк услуг", each Table.RowCount(_), Int64.Type}
    }),
    Reward = Table.AddColumn(Grouped, "Вознаграждение, руб.", each
        if [представитель]="" then 0
        else if [Месяц]="2024-07" then (if [Возрастная категория]="дети" and [#"Оплата за месяц, руб."]>100000 then 5000 else 0)
        else if List.Contains({"2024-08","2024-09"},[Месяц]) then (if [#"Оплата за месяц, руб."]>120000 then 8000 else 0)
        else error "Не заданы правила вознаграждения для нового месяца", Currency.Type),
    Result = Table.ReorderColumns(Reward, {"Месяц","id пациента","представитель","Возрастная категория","Оплата за месяц, руб.","Вознаграждение, руб.","Строк услуг"})
in Table.Sort(Result, {{"Месяц",Order.Ascending},{"id пациента",Order.Ascending}});

shared Вознаграждения =
let
    Agents = Table.SelectRows(ПациентМесяц, each [представитель]<>""),
    Grouped = Table.Group(Agents, {"представитель","Месяц"}, {{"Сумма", each List.Sum([#"Вознаграждение, руб."]), Currency.Type}}),
    Pivot = Table.Pivot(Grouped, {"2024-07","2024-08","2024-09"}, "Месяц", "Сумма", List.Sum),
    Zeros = Table.ReplaceValue(Pivot, null, 0, Replacer.ReplaceValue, {"2024-07","2024-08","2024-09"}),
    Total = Table.AddColumn(Zeros, "Всего, руб.", each [#"2024-07"]+[#"2024-08"]+[#"2024-09"], Currency.Type)
in Table.Sort(Total, {{"представитель",Order.Ascending}});

shared ИтогиМесяцев =
let
    Source = Table.Group(Услуги, {"Месяц"}, {
        {"Строк услуг", each Table.RowCount(_), Int64.Type},
        {"Уникальных пациентов", each List.Count(List.Distinct([id пациента])), Int64.Type},
        {"Компенсация, руб.", each List.Sum([#"компенсация, руб."]), Currency.Type}
    }),
    Rewards = Table.Group(ПациентМесяц, {"Месяц"}, {
        {"Пациентов с выплатой", each Table.RowCount(Table.SelectRows(_, each [#"Вознаграждение, руб."]>0)), Int64.Type},
        {"Вознаграждение, руб.", each List.Sum([#"Вознаграждение, руб."]), Currency.Type}
    }),
    Joined = Table.NestedJoin(Source, {"Месяц"}, Rewards, {"Месяц"}, "Rewards", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Rewards", {"Пациентов с выплатой","Вознаграждение, руб."})
in Table.Sort(Expanded, {{"Месяц",Order.Ascending}});

shared ИтогиПериода = #table({"Показатель за период","Значение"}, {
    {"Строк услуг",Table.RowCount(Услуги)},
    {"Уникальных пациентов",List.Count(List.Distinct(Услуги[id пациента]))},
    {"Компенсация, руб.",List.Sum(Услуги[#"компенсация, руб."])},
    {"Вознаграждение, руб.",List.Sum(ПациентМесяц[#"Вознаграждение, руб."])}
});
