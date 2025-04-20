using System.Text;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Text.RegularExpressions;
using System.Globalization;

using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Models.Bases;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Styles;

namespace RegistrationSummary.Common.Services;

public class ExcelService
{
    private readonly int MAX_ROWS = 200;
    public readonly int ROW_STARTING_INDEX = 6;

    public SheetsService SheetService;
    public string SpreadSheetId => SelectedEvent?.SpreadsheetId ?? string.Empty;
    public Event? SelectedEvent;

    public string SummaryTabName = string.Empty;
    private string _rawDataTabName = string.Empty;
    private string _preprocessedDataTabName = string.Empty;
    private string _groupBalanceTabName = string.Empty;
    private string _leaderText = string.Empty;
    private string _followerText = string.Empty;
    private string _soloText = string.Empty;
    private int[] _prices = Array.Empty<int>();

    public List<Request> Requests = new List<Request>();

    public string? _columnNameForInstallmentsSum = null;
    public string ColumnNameForInstallmentsSum
    {
        get
        {
            if (_columnNameForInstallmentsSum == null)
            {
                var request = SheetService.Spreadsheets.Values.Get(
                    SelectedEvent?.SpreadsheetId,
                    $"{SummaryTabName}!E1"
                );
                var value = request.Execute().Values[0][0].ToString();

                _columnNameForInstallmentsSum = value ?? string.Empty;
            }

            return _columnNameForInstallmentsSum;
        }
    }

    public ExcelService(SheetsService sheetService)
    {
        SheetService = sheetService;
    }

    public void Initialize(Event selectedEvent, string rawDataTabName, string preprocessedDataTabName, string summaryTabName,
        string groupBalanceTabName, string leaderText, string followerText, string soloText, int[] prices)
    {
        SelectedEvent = selectedEvent;
        _rawDataTabName = rawDataTabName;
        _preprocessedDataTabName = preprocessedDataTabName;
        _groupBalanceTabName = groupBalanceTabName;
        _leaderText = leaderText;
        _followerText = followerText;
        _soloText = soloText;
        SummaryTabName = summaryTabName;
        _prices = prices;
    }

    public void SetUpRegistrationsEditableTab()
	{
		var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();
		// Check if {_preprocessedDataTabName} tab is already existing:
		if (spreadsheet.Sheets.Any(sheet => sheet.Properties.Title.Equals(_preprocessedDataTabName)))
			return;

		var preprocessedDataTabNameTmp = _preprocessedDataTabName;
		var sheetId = AddNewSpreadSheet(preprocessedDataTabNameTmp, 26, MAX_ROWS);

		if (sheetId == null)
			return;

		// Headers
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Login}1", $"=\"Login\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Email}1", $"=\"Email\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.FirstName}1", $"=\"First name\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.LastName}1", $"=\"Last name\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.PhoneNumber}1", $"=\"Phone number\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Course}1", $"=\"Course\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Role}1", $"=\"Role\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Partner}1", $"=\"Partner\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Installment}1", $"=\"Installment\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Accepted}1", $"=\"Accepted\"");

        // Add Email Commentarry Columnd after the installment one.
        var acceptedColumnId = ColumnNameToIndex(SelectedEvent.PreprocessedColumns.Accepted) + 6;
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{GetColumnName(++acceptedColumnId)}1", $"=\"Confirmation\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{GetColumnName(++acceptedColumnId)}1", $"=\"Waiting List\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{GetColumnName(++acceptedColumnId)}1", $"=\"Not Enough People\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{GetColumnName(++acceptedColumnId)}1", $"=\"Full Class\"");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{GetColumnName(++acceptedColumnId)}1", $"=\"Missing Partner\"");

        // For aggregated columns there is no method to automatizated solution.
        if (SelectedEvent?.CoursesAreMerged ?? false)
		{
			PopulateRegistrationTabForAggregatedData();
			BatchUpdate();

			return;
		}

		// Non-aggregated data
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Email}2", $"=QUERY({{{_rawDataTabName}!{SelectedEvent?.RawDataColumns.Email}2:{SelectedEvent?.RawDataColumns.Email}}})");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.FirstName}2", $"=QUERY({{{_rawDataTabName}!{SelectedEvent?.RawDataColumns.FirstName}2:{SelectedEvent?.RawDataColumns.FirstName}}})");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.LastName}2", $"=QUERY({{{_rawDataTabName}!{SelectedEvent?.RawDataColumns.LastName}2:{SelectedEvent?.RawDataColumns.LastName}}})");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.PhoneNumber}2", $"=QUERY({{{_rawDataTabName}!{SelectedEvent?.RawDataColumns.PhoneNumber}2:{SelectedEvent?.RawDataColumns.PhoneNumber}}})");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Course}2", $"=ARRAYFORMULA(QUERY(IFERROR(TRIM(REGEXEXTRACT({_rawDataTabName}!{SelectedEvent?.RawDataColumns.Course}2:{SelectedEvent?.RawDataColumns.Course}; \"^[^-]+\")); \"\")))");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Role}2", $"=QUERY({{{_rawDataTabName}!{SelectedEvent?.RawDataColumns.Role}2:{SelectedEvent?.RawDataColumns.Role}}})");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Partner}2", $"=QUERY({{{_rawDataTabName}!{SelectedEvent?.RawDataColumns.Partner}2:{SelectedEvent?.RawDataColumns.Partner}}})");
        AddFormula(sheetId.Value, $"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Installment}2", $"=ARRAYFORMULA(IF(ISBLANK({_rawDataTabName}!{SelectedEvent?.RawDataColumns.Installment}2:{SelectedEvent?.RawDataColumns.Installment}); \"\"; 1))");

		BatchUpdate();

		for (var rowIndex = 2; rowIndex < MAX_ROWS + 2; ++rowIndex)
		{
			// Login column.
			AddFormula(
				sheetId.Value,
				$"{preprocessedDataTabNameTmp}!{SelectedEvent?.PreprocessedColumns.Login}{rowIndex}",
				$"=CONCAT(CONCAT(CONCAT(CONCAT(TRIM(LOWER({SelectedEvent?.PreprocessedColumns.Email}{rowIndex}));\u0022,\u0022);TRIM(LOWER({SelectedEvent?.PreprocessedColumns.FirstName}{rowIndex})));\",\");TRIM(LOWER({SelectedEvent?.PreprocessedColumns.LastName}{rowIndex})))"
			);

			if (rowIndex % 100 == 0)
				BatchUpdate();
		}

		BatchUpdate();
	}

	public void AddInstallmentColumnsForRegistrationTab()
	{
        var sheetId = GetSheetIdForCurrentEvent(_preprocessedDataTabName);

        var acceptedColumnId = ColumnNameToIndex(SelectedEvent.PreprocessedColumns.Accepted) + 1;

        //Add columns for payments from summary tab.
        for (var installmentCounter = 1; installmentCounter <= 5; ++installmentCounter)
        {
			++acceptedColumnId;
            var installmentColumnFromSummary = GetColumnName(ColumnNameToIndex(ColumnNameForInstallmentsSum) + 1 + installmentCounter * 2);

            AddFormula(sheetId.Value, $"{_preprocessedDataTabName}!{GetColumnName(acceptedColumnId)}1", $"=\"Installment {installmentCounter}\"");

            for (var rowInstallment = 2; rowInstallment < MAX_ROWS; ++rowInstallment)
			{                
                AddFormula(sheetId.Value,
                    $"{_preprocessedDataTabName}!{GetColumnName(acceptedColumnId)}{rowInstallment}",
                    $"=ARRAYFORMULA(TEXTJOIN(\", \"; 1; Unique(IF($A{rowInstallment} = Summary!$A${ROW_STARTING_INDEX}:$A; Summary!{installmentColumnFromSummary}${ROW_STARTING_INDEX}:{installmentColumnFromSummary};))))");
            }
        }

		BatchUpdate();
    }

	/// <summary>
	/// If based groups data are aggregated then we need to populate newly come data each time.
	/// </summary>
	public void PopulateRegistrationTabForAggregatedData()
	{
		var request = SheetService.Spreadsheets.Values.Get(
			SelectedEvent?.SpreadsheetId,
			$"{_rawDataTabName}!A2:O"
		);
			
		var rawValues = request.Execute().Values;

		request = SheetService.Spreadsheets.Values.Get(
			SelectedEvent?.SpreadsheetId,
			$"{_preprocessedDataTabName}!A2:O"
		);

		var preprocessedValues = request.Execute().Values;
		var preprocessedSheetId = GetSheetIdForCurrentEvent(_preprocessedDataTabName);
		if (preprocessedSheetId == null)
			return;

		var amountOfAlreadyAddedRows = (preprocessedValues?.Count ?? 0) + 2;

		if ((preprocessedValues?.Count ?? 0) == 0)
		{
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Login}1", $"=\"login\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Email}1", $"=\"Email\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.FirstName}1", $"=\"First name\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.LastName}1", $"=\"Last name\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.PhoneNumber}1", $"=\"Phone number\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Course}1", $"=\"Course\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Role}1", $"=\"Role\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Partner}1", $"=\"Partner\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Installment}1", $"=\"Installment\"");
			AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Accepted}1", $"=\"Accepted\"");
		}

		var textinfo = CultureInfo.CurrentCulture.TextInfo;

		foreach (var value in rawValues)
		{
			var email = value[GetColumnIndex(SelectedEvent?.RawDataColumns.Email) - 1]?.ToString()?.Trim().ToLower();
			var firstName = textinfo.ToTitleCase(value[GetColumnIndex(SelectedEvent?.RawDataColumns.FirstName) - 1]?.ToString()?.Trim() ?? string.Empty);
			var lastName = textinfo.ToTitleCase(value[GetColumnIndex(SelectedEvent?.RawDataColumns.LastName) - 1]?.ToString()?.Trim().ToLower() ?? string.Empty);

			if (email == null || firstName == null || lastName == null)
				continue;

			var courses = value[GetColumnIndex(SelectedEvent?.RawDataColumns.Course) - 1]?.ToString()?.Split(",", StringSplitOptions.RemoveEmptyEntries).Select(name => name.Trim()) ?? null;

			if (courses == null)
				return;

			var login = $"{email},{firstName},{lastName}";

			foreach (var course in courses)
			{
				if (preprocessedValues != null &&
					preprocessedValues.Any(student =>
						(student[GetColumnIndex(SelectedEvent?.PreprocessedColumns.Login) - 1].ToString()?.Equals(login) ?? true) &&
						(student[GetColumnIndex(SelectedEvent?.PreprocessedColumns.Course) - 1].ToString()?.Equals(course) ?? true))
					)
					continue;

				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Login}{amountOfAlreadyAddedRows}", $"=\"{login}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Email}{amountOfAlreadyAddedRows}", $"=\"{email}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.FirstName}{amountOfAlreadyAddedRows}", $"=\"{firstName}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.LastName}{amountOfAlreadyAddedRows}", $"=\"{lastName}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.PhoneNumber}{amountOfAlreadyAddedRows}", $"=\"{value[GetColumnIndex(SelectedEvent?.RawDataColumns.PhoneNumber) - 1]?.ToString()?.Trim()}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Course}{amountOfAlreadyAddedRows}", $"=\"{course.Trim()}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Role}{amountOfAlreadyAddedRows}", $"=\"{value[GetColumnIndex(SelectedEvent?.RawDataColumns.Role) - 1]?.ToString()?.Trim()}\"");
				AddFormula(preprocessedSheetId.Value, $"{_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Partner}{amountOfAlreadyAddedRows}", $"=\"{value[GetColumnIndex(SelectedEvent?.RawDataColumns.Partner) - 1]?.ToString()?.Trim()}\"");
																			 
				amountOfAlreadyAddedRows++;
			}				
		}

		BatchUpdate();
	}

	public void SetUpGroupBalanceTab()
	{
        var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();

        if (spreadsheet.Sheets.Any(sheet => sheet.Properties.Title.Equals(_groupBalanceTabName)))
            return;

        var sheetId = AddNewSpreadSheet(_groupBalanceTabName, 20, 5 + SelectedEvent?.Courses.Count ?? 0);

        if (sheetId == null)
            return;

        AddFormula(sheetId, $"{_groupBalanceTabName}!A4", $"=\"Groups\"");
        AddFormula(sheetId, $"{_groupBalanceTabName}!A5", $"=Sort(Unique({_preprocessedDataTabName}!{SelectedEvent?.PreprocessedColumns.Course}2:{SelectedEvent?.PreprocessedColumns.Course}))");

        AddFormula(sheetId, $"{_groupBalanceTabName}!B1", $"=\"{_leaderText}\"");
        AddFormula(sheetId, $"{_groupBalanceTabName}!C1", $"=\"{_followerText}\"");
        AddFormula(sheetId, $"{_groupBalanceTabName}!D1", $"=\"{_soloText}\"");

        AddFormula(sheetId, $"{_groupBalanceTabName}!B2", $"=\"ALL REGISTRATIONS\"");
		GenerateGroupBalanceSummaryHeader(sheetId, "B");
        AddFormula(sheetId, $"{_groupBalanceTabName}!F2", $"=\"ACCEPTED\"");
        GenerateGroupBalanceSummaryHeader(sheetId, "F");
        AddFormula(sheetId, $"{_groupBalanceTabName}!J2", $"=\"NOT PAID\"");
        GenerateGroupBalanceSummaryHeader(sheetId, "J");
        AddFormula(sheetId, $"{_groupBalanceTabName}!N2", $"=\"PAID\"");
        GenerateGroupBalanceSummaryHeader(sheetId, "N");
        AddFormula(sheetId, $"{_groupBalanceTabName}!R2", $"=\"MISSING\"");
        AddFormula(sheetId, $"{_groupBalanceTabName}!R3", $"=SUM(R5:R)");
        AddFormula(sheetId, $"{_groupBalanceTabName}!S3", $"=SUM(S5:S)");

		BatchUpdate();

        var courseColumn = SelectedEvent?.PreprocessedColumns.Course;
        var roleColumn = SelectedEvent?.PreprocessedColumns.Role;
        var acceptedColumn = SelectedEvent?.PreprocessedColumns.Accepted;
		var firstInstallmentColumn = GetColumnName(GetColumnIndex(SelectedEvent?.PreprocessedColumns.Accepted) + 1);

        for (var rowGroupIndex = 0; rowGroupIndex <= SelectedEvent?.Courses.Count; ++rowGroupIndex)
		{
			var rowIndex = rowGroupIndex + 5;
			// ALL REGISTRATIONS
			AddFormula(sheetId,
                $"{_groupBalanceTabName}!B{rowIndex}", 
				$"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $B$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; \"<>'X'\")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!C{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $C$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; \"<>'X'\")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!D{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $D$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; \"<>'X'\")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!E{rowIndex}",
                $"=SUM(B{rowIndex}:D{rowIndex})");

            // ACCEPTED
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!F{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $B$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; 1)");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!G{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $C$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; 1)");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!H{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $D$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; 1)");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!I{rowIndex}",
                $"=SUM(F{rowIndex}:H{rowIndex})");

            // NOT PAID
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!J{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $B$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; 1; {_preprocessedDataTabName}!${firstInstallmentColumn}$2:${firstInstallmentColumn}; \"\")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!K{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $C$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; 1; {_preprocessedDataTabName}!${firstInstallmentColumn}$2:${firstInstallmentColumn}; \"\")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!L{rowIndex}",
                $"=COUNTIFS({_preprocessedDataTabName}!${courseColumn}$2:${courseColumn}; $A{rowIndex}; {_preprocessedDataTabName}!${roleColumn}$2:${roleColumn}; $D$1; {_preprocessedDataTabName}!${acceptedColumn}$2:${acceptedColumn}; 1; {_preprocessedDataTabName}!${firstInstallmentColumn}$2:${firstInstallmentColumn}; \"\")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!M{rowIndex}",
                $"=SUM(J{rowIndex}:L{rowIndex})");

            // PAID
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!N{rowIndex}",
                $"=F{rowIndex}-J{rowIndex}");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!O{rowIndex}",
                $"=G{rowIndex}-K{rowIndex}");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!P{rowIndex}",
                $"=H{rowIndex}-L{rowIndex}");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!Q{rowIndex}",
                $"=SUM(N{rowIndex}:P{rowIndex})");

            // MISSING
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!R{rowIndex}",
                $"=IF(B{rowIndex}<C{rowIndex};C{rowIndex}-B{rowIndex};\" \")");
            AddFormula(sheetId,
                $"{_groupBalanceTabName}!S{rowIndex}",
                $"=IF(C{rowIndex}<B{rowIndex};B{rowIndex}-C{rowIndex};\" \")");
        }

        Requests.AddRange(
			GetGroupBalanceFormattingRequests(sheetId.Value, 5 + SelectedEvent?.Courses.Count ?? 0));

        BatchUpdate();
    }

    public IList<Request> GetGroupBalanceFormattingRequests(int sheetId, int rowCount)
    {
        var requests = new List<Request>();

        // Merge category header blocks
        requests.AddRange(new[]
        {
			MergeCells(sheetId, 1, 2, 1, 5),   // B-E
			MergeCells(sheetId, 1, 2, 5, 9),   // F-I
			MergeCells(sheetId, 1, 2, 9, 13),  // J-M
			MergeCells(sheetId, 1, 2, 13, 17), // N-Q
			MergeCells(sheetId, 1, 2, 17, 19)  // R-S
		});

        // Bold + wrap + center all headers except B1,C1
        requests.Add(new Request
        {
            RepeatCell = new RepeatCellRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId,
                    StartRowIndex = 1,
                    EndRowIndex = 4
                },
                Cell = new CellData { UserEnteredFormat = SharedFormats.BoldCenterWrapHeaderFormat.Format },
                Fields = SharedFormats.BoldCenterWrapHeaderFormat.Fields
            }
        });

        // Set borders: under header row (1), under totals (3)
        requests.AddRange(new[]
        {
			BorderRow(sheetId, 1), // Under category row
			BorderRow(sheetId, 3)  // Under sum row
		});

        // Vertical separators
        foreach (var col in new[] { 0, 4, 8, 12, 16 })
        {
            requests.Add(BorderColumn(sheetId, col, rowCount, 2));
        }

        // Narrow all columns B–S
        requests.Add(new Request
        {
            UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
            {
                Range = new DimensionRange
                {
                    SheetId = sheetId,
                    Dimension = "COLUMNS",
                    StartIndex = 1,
                    EndIndex = 19
                },
                Properties = SharedFormats.NarrowColumnWidth.Props,
                Fields = SharedFormats.NarrowColumnWidth.Fields
            }
        });

        return requests;
    }

    private Request MergeCells(int sheetId, int startRow, int endRow, int startCol, int endCol) => new()
    {
        MergeCells = new MergeCellsRequest
        {
            Range = new GridRange
            {
                SheetId = sheetId,
                StartRowIndex = startRow,
                EndRowIndex = endRow,
                StartColumnIndex = startCol,
                EndColumnIndex = endCol
            },
            MergeType = "MERGE_ALL"
        }
    };

    private Request BorderRow(int sheetId, int rowIndex) => new()
    {
        UpdateBorders = new UpdateBordersRequest
        {
            Range = new GridRange
            {
                SheetId = sheetId,
                StartRowIndex = rowIndex,
                EndRowIndex = rowIndex + 1,
                StartColumnIndex = 0,
                EndColumnIndex = 19
            },
            Bottom = SharedFormats.SolidBlackBorder
        }
    };

    private Request BorderColumn(int sheetId, int columnIndex, int totalRows, int startRowIndex = 0) => new()
    {
        UpdateBorders = new UpdateBordersRequest
        {
            Range = new GridRange
            {
                SheetId = sheetId,
                StartRowIndex = startRowIndex,
                EndRowIndex = totalRows,
                StartColumnIndex = columnIndex,
                EndColumnIndex = columnIndex + 1
            },
            Right = SharedFormats.SolidBlackBorder
        }
    };

    private void GenerateGroupBalanceSummaryHeader(int? sheetId, string columnName)
	{
		var columnIndex = GetColumnIndex(columnName);

		for (var columnCounter = 0; columnCounter < 4; ++columnCounter)
		{
			string columnGenerationName = GetColumnName(columnIndex + columnCounter);
            AddFormula(sheetId, $"{_groupBalanceTabName}!{columnGenerationName}3", $"=SUM({columnGenerationName}5:{columnGenerationName})");

			var text = "";
			switch(columnCounter)
			{
				case 0:
					text = "Lead";
					break;
                case 1:
                    text = "Follow";
                    break;
                case 2:
                    text = "Solo";
                    break;
                case 3:
                    text = "Sum";
                    break;
            }

            AddFormula(sheetId, $"{_groupBalanceTabName}!{columnGenerationName}4", $"=\"{text}\"");
        }
	}

	public void SetUpSummaryTab()
	{
		var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();

		if (spreadsheet.Sheets.Any(sheet => sheet.Properties.Title.Equals(SummaryTabName)))
			return;

		var sheetId = AddNewSpreadSheet(SummaryTabName, 30 + SelectedEvent?.Courses.Count ?? 0, MAX_ROWS);

		if (sheetId == null)
			return;

		HideColumn(sheetId.Value, "A");
		AddFormula(sheetId.Value, $"{SummaryTabName}!A{ROW_STARTING_INDEX}", $"=Unique({_preprocessedDataTabName}!A2:A9999)");

		AddFormula(sheetId.Value, $"{SummaryTabName}!B1", $"={_prices[0]}");
		AddFormula(sheetId.Value, $"{SummaryTabName}!C1", $"={_prices[1]}");
		AddFormula(sheetId.Value, $"{SummaryTabName}!D1", $"={_prices[2]}");
			
		if (_prices.Count() > 3)
			AddFormula(sheetId.Value, $"{SummaryTabName}!F1", $"={_prices[3]}");

		BatchUpdate();

		// Courses headers.
		var courseColumnIndex = 2;
		foreach (var product in SelectedEvent?.Courses)
		{
			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(courseColumnIndex)}2",
				$"=\"{product.Name}\""
			);
			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(courseColumnIndex)}3",
				$"=\"{(((Course)product).IsShorter ? 1 : 0)}\""
			);

			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(courseColumnIndex)}4",
				$"=\"{StackLetters(((Course)product).Code)}\""
			);
			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(courseColumnIndex)}5",
				$"=SUMA({GetColumnName(courseColumnIndex)}{ROW_STARTING_INDEX}:{GetColumnName(courseColumnIndex)}{MAX_ROWS + ROW_STARTING_INDEX})"
			);
			SetColumnWidth(sheetId.Value, courseColumnIndex, 21);
			++courseColumnIndex;
		}
		HideRow(sheetId.Value, 1);
		HideRow(sheetId.Value, 2);

		BatchUpdate();

		var headersColumnIndex = courseColumnIndex;

		var fullCoursesCountColumnIndex = headersColumnIndex;
        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex)}{ROW_STARTING_INDEX - 2}", $"=SUM({GetColumnName(headersColumnIndex)}{ROW_STARTING_INDEX}:{GetColumnName(headersColumnIndex)})");
        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex)}{ROW_STARTING_INDEX - 1}", $"=\"{StackLetters("Sum Full")}\"");
		SetColumnWidth(sheetId.Value, headersColumnIndex++, 21);

		if (_prices.Count() > 3)
		{
            AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex)}{ROW_STARTING_INDEX - 2}", $"=SUM({GetColumnName(headersColumnIndex)}{ROW_STARTING_INDEX}:{GetColumnName(headersColumnIndex)})");
            AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex)}{ROW_STARTING_INDEX - 1}", $"=\"{StackLetters("Sum Cheaper")}\"");
			SetColumnWidth(sheetId.Value, headersColumnIndex++, 21);
        }

        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"Email\"");
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"First Name\"");
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"Last Name\"");
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"Phone Number\"");
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"Courses\"");
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"Partner\"");

		var installmentNeededColumnName = GetColumnName(headersColumnIndex);
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", "=\"Installment\"");

		var paymentAmountColumnIndex = headersColumnIndex;

		var discountColumn = headersColumnIndex;
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"{"Discount"}\"");
		var installmentSumColumn = headersColumnIndex;
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"{"Installment\nSum"}\"");
		var needToBePaidColumn = headersColumnIndex;
		AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"{"Need to\nbe Paid"}\"");

		var summaryCostColumn = headersColumnIndex;

		var installmentColumns = new List<int>();
		AddFormula(sheetId.Value, $"{SummaryTabName}!E1", $"=\"{GetColumnName(installmentSumColumn)}\"");
		// Save informtion in fixed cell in which column installment's columns starts.

		// Installment columns
		for (var i = 0; i < 5; i++)
		{
			installmentColumns.Add(headersColumnIndex);
			AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Inst. {i + 1}\nAmount\"");
			AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Inst. {i + 1}\nDate\"");
		}

        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Confirmation\"");
        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Waiting List\"");
        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Not Enough People\"");
        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Full Class\"");
        AddFormula(sheetId.Value, $"{SummaryTabName}!{GetColumnName(headersColumnIndex++)}{ROW_STARTING_INDEX - 1}", $"=\"Missing Partner\"");

        BatchUpdate();

		// Cost per student
		for (
			var rowIndex = ROW_STARTING_INDEX;
			rowIndex < MAX_ROWS + ROW_STARTING_INDEX;
			++rowIndex
		)
		{
			var cellAmountOfFullCourses = GetColumnName(fullCoursesCountColumnIndex) + rowIndex;

			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{cellAmountOfFullCourses}",
				$"=SUMIF($B$3:${GetColumnName(fullCoursesCountColumnIndex - 1)}$3;\"=0\";$B{rowIndex}:${GetColumnName(fullCoursesCountColumnIndex - 1)}{rowIndex})"
			);

			if (_prices.Count() > 3)
			{
				var shorterCoursesCountColumnIndex = fullCoursesCountColumnIndex + 1;
				var cellAmountOfShorterCourses = GetColumnName(shorterCoursesCountColumnIndex) + rowIndex;
            AddFormula(
                sheetId.Value,
                $"{SummaryTabName}!{cellAmountOfShorterCourses}",
                $"=SUMIF($B$3:${GetColumnName(shorterCoursesCountColumnIndex - 1)}$3;\"=1\";$B{rowIndex}:${GetColumnName(shorterCoursesCountColumnIndex - 1)}{rowIndex})"
            );
        }

			// Cost per student
			SetCellFormat(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(discountColumn)}{rowIndex}",
                SharedFormats.CurrencyPlnCellFormat.Format,
                SharedFormats.CurrencyPlnCellFormat.Fields
			);

			// Sum to be paid for all courses minus discount.
			SetCellFormat(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(installmentSumColumn)}{rowIndex}",
                SharedFormats.CurrencyPlnCellFormat.Format,
                SharedFormats.CurrencyPlnCellFormat.Fields
			);

			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(installmentSumColumn)}{rowIndex}",
				$"=IF({cellAmountOfFullCourses}=1;$B$1;IF({cellAmountOfFullCourses}=2;$B$1+$C$1;IF({cellAmountOfFullCourses}>2;$B$1+$C$1+$D$1*({cellAmountOfFullCourses}-2);0))) - {GetColumnName(discountColumn)}{rowIndex} + IF(AND({cellAmountOfFullCourses}> 0;{installmentNeededColumnName}{rowIndex}<>\"\");20;0)"
			);

			// Rest amount of money to be paid minus already paid installments.
			SetCellFormat(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(needToBePaidColumn)}{rowIndex}",
                SharedFormats.CurrencyPlnCellFormat.Format,
                SharedFormats.CurrencyPlnCellFormat.Fields
			);

			AddFormula(
				sheetId.Value,
				$"{SummaryTabName}!{GetColumnName(needToBePaidColumn)}{rowIndex}",
				$"={GetColumnName(installmentSumColumn)}{rowIndex}-("
				+ installmentColumns.Select(installment => $"{GetColumnName(installment)}{rowIndex}")
					.Aggregate(
						(current, next) =>
						$"{current} + {next}")
				+")"
			);

			// Matching courses to students.
			for (
				var referenceColumnIndex = 2;
				referenceColumnIndex < courseColumnIndex;
				++referenceColumnIndex
			)
			{
				//=IFERROR(INDEX(Registrations!$O$2:$O$100; MATCH(1; (Registrations!$A$2:$A$100=$A5) * (Registrations!$H$2:$H$100=G$2); 0)); "")
				AddFormula(
					sheetId.Value,
					$"{SummaryTabName}!{GetColumnName(referenceColumnIndex)}{rowIndex}",
					$"=IFERROR(INDEX({_preprocessedDataTabName}!${SelectedEvent?.PreprocessedColumns.Accepted}$2:${SelectedEvent?.PreprocessedColumns.Accepted}${MAX_ROWS}; MATCH(1; ({_preprocessedDataTabName}!${SelectedEvent?.PreprocessedColumns.Login}$2:${SelectedEvent?.PreprocessedColumns.Login}${MAX_ROWS}=$A{rowIndex}) *({_preprocessedDataTabName}!${SelectedEvent?.PreprocessedColumns.Course}$2:${SelectedEvent?.PreprocessedColumns.Course}${MAX_ROWS} = {GetColumnName(referenceColumnIndex)}$2); 0)); \"\")"
				);
			}

			var studentColumnIndex = _prices.Count() > 3 ? fullCoursesCountColumnIndex + 2 : fullCoursesCountColumnIndex + 1;

			// Adding students informations.
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.Email, studentColumnIndex++);
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.FirstName, studentColumnIndex++);
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.LastName, studentColumnIndex++);
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.PhoneNumber, studentColumnIndex++);
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.Course, studentColumnIndex++);
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.Partner, studentColumnIndex++);
			AddStudentInformation(sheetId.Value, rowIndex, SelectedEvent?.PreprocessedColumns.Installment, studentColumnIndex++);

			if (rowIndex % 50 == 0)
				BatchUpdate();
		}

		BatchUpdate();

		AddInstallmentColumnsForRegistrationTab();
	}

    public void AddStudentInformation(int sheetId, int rowIndex, string? referenceColumn, int targetColumnIndex)
	{
		if (referenceColumn == null)
			return;

		AddFormula(
			sheetId,
			$"{SummaryTabName}!{GetColumnName(targetColumnIndex)}{rowIndex}",
			$"=ARRAYFORMULA(TEXTJOIN(\", \";1;Unique(JEŻELI($A{rowIndex}={_preprocessedDataTabName}!$A$2:$A;{_preprocessedDataTabName}!{referenceColumn}$2:{referenceColumn};))))"
		);
	}

	public void SetUpCourseTab(Course course)
	{
		var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();

		// Check if {_preprocessedDataTabName} tab is already existing:
		if (
			spreadsheet.Sheets.Any(
				sheet => sheet.Properties.Title.Equals(course.Code)
			)
		)
			return;

		var sheetId = AddNewSpreadSheet(course.Code);

		if (sheetId == null)
			return;

		AddFormula(sheetId.Value, $"{course.Code}!A2", $"=\"Wpłata\"");

		for (var columnIndex = 0; columnIndex < 10; ++columnIndex)
		{
			var columnName = (char)('B' + columnIndex);
			AddFormula(
				sheetId.Value,
				$"{course.Code}!{columnName}1",
				$"=\"{course.Start.AddDays(columnIndex * 7).ToString("dd.MM")}\""
			);
			SetColumnWidth(sheetId.Value, ColumnNameToIndex(columnName.ToString()), 40);
			AddFormula(
				sheetId.Value,
				$"{course.Code}!{columnName}2",
				$"=\"{columnIndex + 1}\""
			);
		}

		AddFormula(
			sheetId.Value,
			$"{course.Code}!L1",
			$"=\"{course.Name}\""
		);
		HideColumn(sheetId.Value, "L");
		HideColumn(sheetId.Value, "Q");

		AddFormula(
			sheetId.Value,
			$"{course.Code}!L2",
			$"=QUERY({{{_preprocessedDataTabName}!C1:O}};CONCATENATE(\"Select Col2, Col3, Col4, Col5 where Col2 is not null and LOWER(Col6) like '%\";LOWER(L1);\"%' and Col13=1\"); 1)"
		);

		BatchUpdate();
	}

	public List<Student> GetStudentsFromRegularSemestersSheet()
	{
		var request = SheetService.Spreadsheets.Values.Get(
			SelectedEvent?.SpreadsheetId,
			$"{SummaryTabName}!A{ROW_STARTING_INDEX}:BZ"
		);
		var values = request.Execute().Values;
		var students = new List<Student>();

		if (values == null || values.Count == 0)
			return new List<Student>();

		var courseColumns = GetColumnsForCoursesInSummaryTab();
		var headersColumns = GetColumnsHeadersInSummaryTab();
		var installmentSumColumn = ColumnNameForInstallmentsSum;
		var confirmationEmailSentColumn = GetColumnIndex(installmentSumColumn) + 11;
		var waitingListEmailSentColumn = confirmationEmailSentColumn + 1;
		var notEnoughPeopleEmailSentColumn = waitingListEmailSentColumn + 1;
		var fullClassEmailSentColumn = notEnoughPeopleEmailSentColumn + 1;
		var missingPartnerEmailSentColumn = fullClassEmailSentColumn + 1;

		var emailColumnIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Email")).ColumnName);
		var firstNameColumnIndex =  ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("First Name")).ColumnName);
		var lastNameColumnIndex =  ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Last Name")).ColumnName);
		var installmentColumnIndex =  ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Installment")).ColumnName);
		var coursesColumnIndex =  ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Courses")).ColumnName);
		var rowId = 0;

		foreach (var row in values)
		{				
			var email = row[emailColumnIndex]?.ToString()?.Trim();

			if (string.IsNullOrEmpty(email))
				continue;

			var newStudent =
				new Student
				{
					Id = rowId++,
					Email = email,
					FirstName = row[firstNameColumnIndex]?.ToString()?.Trim() ?? string.Empty,
					LastName = row[lastNameColumnIndex]?.ToString()?.Trim() ?? string.Empty,
					PaymentAmount = int.Parse(row[ColumnNameToIndex(installmentSumColumn) + 1]?.ToString()?.Replace(" zł", "")?.Replace(" ", "") ?? "0"),
					Installments = (row[installmentColumnIndex]?.ToString()?.Trim() ?? string.Empty).Equals("1"),
					Courses = SelectedEvent?.Courses
						.Where(
							course =>
								(row[coursesColumnIndex]?.ToString() ?? string.Empty)										
									.Split(",", StringSplitOptions.RemoveEmptyEntries)
									.Select(name => name.Trim())
									.Any(
										name =>
											name.Contains(
												course.Name,
												StringComparison.OrdinalIgnoreCase
											)
									)
						)
						.Select(product => (Course)product)
						.Select(course => new Course()
						{
							Type = course.Type,
							Name = course.Name,
							Code = course.Code,
							Start = course.Start,
							End = course.End,
							DayOfWeek = course.DayOfWeek,
							Time = course.Time,
							Location = course.Location,
							AdditionalComment = course.AdditionalComment,
							IsSolo = course.IsSolo
						})
						.ToList()
				};

			if (row.Count > confirmationEmailSentColumn && !string.IsNullOrEmpty(row[confirmationEmailSentColumn].ToString()))
			{
				newStudent.AlreadySentEmails.Add(EmailType.Confirmation);
			}

			if (row.Count > waitingListEmailSentColumn && !string.IsNullOrEmpty(row[waitingListEmailSentColumn].ToString()))
			{
				newStudent.AlreadySentEmails.Add(EmailType.WaitingList);
			}

			if (row.Count > notEnoughPeopleEmailSentColumn && !string.IsNullOrEmpty(row[notEnoughPeopleEmailSentColumn].ToString()))
			{
				newStudent.AlreadySentEmails.Add(EmailType.NotEnoughPeople);
			}

			if (row.Count > fullClassEmailSentColumn && !string.IsNullOrEmpty(row[fullClassEmailSentColumn].ToString()))
			{
				newStudent.AlreadySentEmails.Add(EmailType.FullClass);
			}

			if (row.Count > missingPartnerEmailSentColumn && !string.IsNullOrEmpty(row[missingPartnerEmailSentColumn].ToString()))
			{
				newStudent.AlreadySentEmails.Add(EmailType.MissingPartner);
			}

			foreach (var course in courseColumns)
			{
				var foundCourse =
					newStudent
						.Courses?
						.SingleOrDefault(studentCourse => studentCourse.Code.Equals(course.CourseCode));

				if (foundCourse == null)
					continue;

				var value = row[ColumnNameToIndex(course.ColumnName)]?.ToString() ?? string.Empty;

				switch (value.ToLower())
				{
					case "1":
						foundCourse.Status = EmailType.Confirmation;
						break;

					case "w":
						foundCourse.Status = EmailType.WaitingList;
						break;

					case "nep":
						foundCourse.Status = EmailType.NotEnoughPeople;
						break;

					case "fc":
						foundCourse.Status = EmailType.FullClass;
						break;

					case "bp":
						foundCourse.Status = EmailType.MissingPartner;
						break;
				}
			}

			students.Add(newStudent);
		}

		students = students
			.Where(student => !string.IsNullOrEmpty(student.Email))
			.Where(student => (student?.Courses?.Count ?? 0) > 0)
			.ToList();

        GetAdditionalCommentaryForEmail(students);

		return students;
	}

	private void GetAdditionalCommentaryForEmail(List<Student> students)
	{
		var request = SheetService.Spreadsheets.Values.Get(
			SelectedEvent?.SpreadsheetId,
			$"{_preprocessedDataTabName}!A{2}:BZ"
		);

		var headersColumns = GetColumnsHeadersInTab(_preprocessedDataTabName);
		var loginColumnIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("login")).ColumnName);
		var courseColumnIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Course")).ColumnName);
		var acceptedColumnIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Accepted")).ColumnName);

		var loginConfirmationCommentaryIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Confirmation")).ColumnName);
		var loginWaitingListCommentaryIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Waiting List")).ColumnName);
		var loginNotEnoughPeopleCommentaryIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Not Enough People")).ColumnName);
		var loginFullClassCommentaryIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Full Class")).ColumnName);
		var loginMissingPartnerCommentaryIndex = ColumnNameToIndex(headersColumns.Single(header => header.Header.Equals("Missing Partner")).ColumnName);

		var preprocessedData = request.Execute().Values
            .Where(r => r.Count > loginColumnIndex && !string.IsNullOrEmpty(r[loginColumnIndex]?.ToString()))
            .Where(r => r.Count > courseColumnIndex && !string.IsNullOrEmpty(r[courseColumnIndex]?.ToString()))
            .Where(r => r.Count > acceptedColumnIndex && !string.IsNullOrEmpty(r[acceptedColumnIndex]?.ToString()))
			.ToList();

        var studentsData = new List<(string? Login, string? Course, string? Accepted, string? Confirmation, string? WaitingList, string? NotEnoughPeople, string? FullClass, string? MissingPartner)>();
        foreach (var row in preprocessedData)
        {
            studentsData.Add((
            Login: row[loginColumnIndex].ToString(),
            Course: row[courseColumnIndex].ToString(),
            Accepted: row[acceptedColumnIndex].ToString(),
			Confirmation: row.Count > loginConfirmationCommentaryIndex ? row[loginConfirmationCommentaryIndex]?.ToString() : null,
			WaitingList: row.Count > loginWaitingListCommentaryIndex ? row[loginWaitingListCommentaryIndex]?.ToString() : null,
			NotEnoughPeople: row.Count > loginNotEnoughPeopleCommentaryIndex ? row[loginNotEnoughPeopleCommentaryIndex]?.ToString() : null,
			FullClass: row.Count > loginFullClassCommentaryIndex ? row[loginFullClassCommentaryIndex]?.ToString() : null,
			MissingPartner: row.Count > loginMissingPartnerCommentaryIndex ? row[loginMissingPartnerCommentaryIndex]?.ToString() : null
            ));
        }

        foreach (var student in students)
		{
            foreach (var course in student.Courses)
			{
                string commentary = string.Empty;
				var foundStudent = studentsData
					.FirstOrDefault(stu => stu.Login.ToLower().Equals(student.Login.ToLower()) && stu.Course.ToLower().Equals(course.Name.ToLower()));

				if (foundStudent.Login == null)
					continue;

                switch (course.Status)
                {
                    case EmailType.Confirmation:
                        commentary = foundStudent.Confirmation;
                        break;
                    case EmailType.WaitingList:
                        commentary = foundStudent.WaitingList;
                        break;
                    case EmailType.NotEnoughPeople:
                        commentary = foundStudent.NotEnoughPeople;
                        break;
                    case EmailType.FullClass:
                        commentary = foundStudent.FullClass;
                        break;
                    case EmailType.MissingPartner:
                        commentary = foundStudent.MissingPartner;
                        break;
                    default:
                        return;
                }

				course.EmailCommentary = commentary;
            }
        }
    }

	private List<(string CourseCode, string ColumnName)> GetColumnsForCoursesInSummaryTab()
	{
		var request = SheetService.Spreadsheets.Values.Get(
			SelectedEvent?.SpreadsheetId,
			$"{SummaryTabName}!B{ROW_STARTING_INDEX - 2}:AZ{ROW_STARTING_INDEX - 2}"
		);

		var values = request.Execute().Values;


		var list = new List<(string CourseShortcut, string ColumnName)>();

		if (values == null || values.Count == 0)
			return list;

		var counter = 2; // 'B'
		foreach (var cell in values[0])
		{
			var value = cell?.ToString()?.Replace("\n", "") ?? string.Empty;

			if (value.ToLower().Equals("sum") || string.IsNullOrEmpty(value))
				break;

			list.Add((value, GetColumnName(counter++).ToString()));
		}

		return list;
	}

	private List<(string Header, string ColumnName)> GetColumnsHeadersInSummaryTab()
	{
		var request = SheetService.Spreadsheets.Values.Get(
			SelectedEvent?.SpreadsheetId,
			$"{SummaryTabName}!B{ROW_STARTING_INDEX - 1}:BZ{ROW_STARTING_INDEX - 1}"
		);

		var values = request.Execute().Values;

		var list = new List<(string Header, string ColumnName)>();

		if (values == null || values.Count == 0)
			return list;

		var counter = 2; // 'B'
		var headersStarted = false;

		foreach (var cell in values[0])
		{
			var value = cell.ToString();

			if (!headersStarted)
			{
				if (!value?.Equals("Email") ?? true)
				{
					counter++;
					continue;
				}

				headersStarted = true;
			}

			list.Add((value ?? string.Empty, GetColumnName(counter++).ToString()));

			if (value?.ToLower().Equals("Installment") ?? false)
				break;
		}

		return list;
	}

    private List<(string Header, string ColumnName)> GetColumnsHeadersInTab(string tabName, int startingRowIndex = 1)
    {
        var request = SheetService.Spreadsheets.Values.Get(
            SelectedEvent?.SpreadsheetId,
            $"{tabName}!A{startingRowIndex}:BZ{startingRowIndex}"
        );

        var values = request.Execute().Values;

        var list = new List<(string Header, string ColumnName)>();

        if (values == null || values.Count == 0)
            return list;

        var counter = 1;

        foreach (var cell in values[0])
        {
            var value = cell.ToString();

			if (string.IsNullOrEmpty(value))
				continue;
			
			list.Add((value ?? string.Empty, GetColumnName(counter++).ToString()));
        }

        return list;
    }

    public void SetUpTabForAccountants()
	{
		var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();
    var sheetName = "Accounting";

    // Check if {_preprocessedDataTabName} tab is already existing:
    if (
			spreadsheet.Sheets.Any(
				sheet => sheet.Properties.Title.Equals(sheetName)
			)
		)
			return;

		var sheetId = AddNewSpreadSheet(sheetName);

		if (sheetId == null)
			return;

		AddFormula(sheetId.Value, $"{sheetName}!A1", $"=\"Month:\"");
    AddFormula(sheetId.Value, $"{sheetName}!B1", $"=\"{DateTime.Today.ToString("yyyy-MM-dd")}\"");

    AddFormula(sheetId.Value, $"{sheetName}!A2", $"=\"First Name\"");
		AddFormula(sheetId.Value, $"{sheetName}!B2", $"=\"Last Name\"");
		AddFormula(sheetId.Value, $"{sheetName}!C2", $"=\"Amount\"");
		AddFormula(sheetId.Value, $"{sheetName}!D2", $"=\"Date\"");

    var emailColumn = GetColumnsHeadersInSummaryTab().Single(column => column.Header.Equals("Email")).ColumnName;
    var installmentSumColumn = ColumnNameForInstallmentsSum;

		var lastColumnNeeded = GetColumnIndex(installmentSumColumn) + 11; // 5 * 2 installment columns

		var subqueries = string.Empty;
		for (var i = 0; i < 5; ++i)
		{
			var amountColumn = (11 + i * 2).ToString();
			var dateColumn = (12 + i * 2).ToString();

			subqueries +=
				$";QUERY({{Summary!{emailColumn}{ROW_STARTING_INDEX}:{GetColumnName(lastColumnNeeded)}}};\"Select Col1, Col2, Col{amountColumn}, Col{dateColumn}\")";
		}

		subqueries = $"=QUERY({{{subqueries.Substring(1)}}};\"SELECT * Where Col3 is not null And MONTH(Col4)+1 = \" & MONTH(B1) & \" And YEAR(Col4) = \" & YEAR(B1))";

		AddFormula(sheetId.Value,
			$"{sheetName}!A3",
			subqueries);

		BatchUpdate();
	}

    public void SetUpTabForNoPayments()
    {
        var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();
			var sheetName = "NoPayments";

        // Check if {_preprocessedDataTabName} tab is already existing:
        if (
            spreadsheet.Sheets.Any(
                sheet => sheet.Properties.Title.Equals(sheetName)
            )
        )
            return;

        var sheetId = AddNewSpreadSheet(sheetName);

        if (sheetId == null)
            return;

        AddFormula(sheetId.Value, $"{sheetName}!A1", $"=\"Shows people who had got an confirmation email more than 7 days ago and did not pay yet.\"");
        AddFormula(sheetId.Value, $"{sheetName}!E1", $"=\"Today\"");
        AddFormula(sheetId.Value, $"{sheetName}!F1", $"=7");

        AddFormula(sheetId.Value, $"{sheetName}!A2", $"=\"Email\"");
        AddFormula(sheetId.Value, $"{sheetName}!B2", $"=\"First Name\"");
        AddFormula(sheetId.Value, $"{sheetName}!C2", $"=\"Last Name\"");
        AddFormula(sheetId.Value, $"{sheetName}!D2", $"=\"Confirmation Sent\"");
        AddFormula(sheetId.Value, $"{sheetName}!E2", $"=\"Confirmation Date\"");
        AddFormula(sheetId.Value, $"{sheetName}!F2", $"=\"Older than \" & F1 & \" days\"");

			var emailColumn = GetColumnsHeadersInSummaryTab().Single(column => column.Header.Equals("Email")).ColumnName;
			var confirmationEmailColumn = GetColumnName(GetColumnIndex(ColumnNameForInstallmentsSum) + 12);

        AddFormula(sheetId.Value, $"{sheetName}!A3", 
				$"=QUERY(" +
				$"	{{QUERY({{Summary!{emailColumn}{ROW_STARTING_INDEX}:{confirmationEmailColumn}}}; \"Select Col1, Col2, Col3, Col12, Col14, Col16, Col18, Col21\")}};" +
				$"\"SELECT Col1, Col2, Col3, Col8 " +
				$"Where Col1 is not null " +
				$"And Col4 is null " +
				$"And Col5 is null " +
				$"And Col6 is null " +
				$"And Col7 is null " +
				$"And Col8 is not null\")");

			for (int i = 3; i < 200; ++i)
			{
				AddFormula(sheetId.Value, $"{sheetName}!E{i}", $"=LEFT(D{i};10)");
            AddFormula(sheetId.Value, $"{sheetName}!F{i}", $"=IF(AND(E{i}<>\"\"; TODAY() - E{i} > $F$1); 1; \" \")");
        }

        BatchUpdate();
    }

    public string StackLetters(string input)
	{
		if (string.IsNullOrEmpty(input))
			return input;

		var result = new StringBuilder();
		for (int i = 0; i < input.Length - 1; i++)
		{
			result.Append(input[i]);
			result.Append('\n'); // Add a newline character between each letter
		}
		result.Append(input[input.Length - 1]);

		return result.ToString();
	}

	public int? AddNewSpreadSheet(string sheetName, int numColumns = 26, int numRows = 1000)
	{
		var addSheetRequest = new Request()
		{
			AddSheet = new AddSheetRequest()
			{
				Properties = new SheetProperties()
				{
					Title = sheetName,
					GridProperties = new GridProperties()
					{
						ColumnCount = numColumns,
						RowCount = numRows
					}
				}
			}
		};

		// Add the new sheet to the spreadsheet
		var batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
		batchUpdateRequest.Requests = new List<Request>() { addSheetRequest };

		bool success = false;
		int retries = 0;
		int maxRetries = 10; // Maximum number of retries

		BatchUpdateSpreadsheetResponse? batchUpdateResponse = null;

		while (!success && retries < maxRetries)
		{
			try
			{
				batchUpdateResponse =
					SheetService.Spreadsheets
						.BatchUpdate(batchUpdateRequest, SpreadSheetId)
						.Execute();

				success = true; // Request succeeded, exit loop
			}
			catch (Google.GoogleApiException ex)
			{
				if (ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests)
				{
					Console.WriteLine("Quota exceeded. Retrying in 30 seconds...");
					retries++;
					Thread.Sleep(30000); // Wait for 30 seconds before retrying
				}
				else
				{
					// Handle other exceptions
					Console.WriteLine("An error occurred: " + ex.Message);
					break;
				}
			}
			catch
			{
				// FUCK THIS SHIT
			}
		}

		return batchUpdateResponse?.Replies[0]?.AddSheet?.Properties?.SheetId;
	}

	public void AddFormula(int? sheetId, string cell, string formula)
	{
		if (sheetId == null || sheetId <= 0)
			return;

		var columnName = cell.Split('!')[1].TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

		var updateRequest = new Request
		{
			UpdateCells = new UpdateCellsRequest
			{
				Start = new GridCoordinate
				{
					SheetId = sheetId,
					RowIndex = Convert.ToInt32(cell.Split('!')[1].Substring(columnName.Length)) - 1,
					ColumnIndex = GetColumnIndex(columnName) - 1
				},
				Rows = new List<RowData>
				{
					new RowData
					{
						Values = new List<CellData>
						{
							new CellData
							{
								UserEnteredValue = new ExtendedValue { FormulaValue = formula },
							}
						}
					}
				},
				Fields = "userEnteredValue.formulaValue"
			}
		};

		Requests.Add(updateRequest);
	}

	public void SetCellFormat(
		int sheetId,
		string cell,
		CellFormat? cellFormat = null,
		string? fieldsFormat = null)
	{
		var rowIndex = GetRowFromCell(cell);
		var columnIndex = GetColumnIndex(GetColumnFromCell(cell));

		// Create a new request to update the number format
		var request = new Request
		{
			RepeatCell = new RepeatCellRequest
			{
				Range = new GridRange
				{
					SheetId = sheetId,
					StartRowIndex = rowIndex - 1,
					EndRowIndex = rowIndex,
					StartColumnIndex = columnIndex,
					EndColumnIndex = columnIndex + 1
				},
				Cell = new CellData { UserEnteredFormat = cellFormat },
				Fields = "userEnteredFormat.numberFormat"
			}
		};

		Requests.Add(request);
	}

	private string GetColumnFromCell(string cell)
	{
		string pattern = @"([A-Za-z]+)(\d+)";
		Match match = Regex.Match(cell, pattern);

		return match.Groups[1].Value;
	}

	private int GetRowFromCell(string cell)
	{
		string pattern = @"([A-Za-z]+)(\d+)";
		Match match = Regex.Match(cell, pattern);

		return int.Parse(match.Groups[2].Value);
	}

	public void HideRow(int sheetId, int rowIndex)
	{
		var hideColumnRequest = new Request
		{
			UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
			{
				Range = new DimensionRange
				{
					SheetId = sheetId,
					Dimension = "ROWS",
					StartIndex = rowIndex - 1,
					EndIndex = rowIndex
				},
				Properties = new DimensionProperties { HiddenByUser = true },
				Fields = "hiddenByUser"
			}
		};

		Requests.Add(hideColumnRequest);
	}

	public void HideColumn(int sheetId, string columnName)
	{
		var hideColumnRequest = new Request
		{
			UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
			{
				Range = new DimensionRange
				{
					SheetId = sheetId,
					Dimension = "COLUMNS",
					StartIndex = ColumnNameToIndex(columnName),
					EndIndex = ColumnNameToIndex(columnName) + 1
				},
				Properties = new DimensionProperties { HiddenByUser = true },
				Fields = "hiddenByUser"
			}
		};

		Requests.Add(hideColumnRequest);
	}

	private int ColumnNameToIndex(string columnName)
	{
		if (columnName.Length == 1)
		{
			return columnName[0] - 'A';
		}
		else if (columnName.Length == 2)
		{
			return (columnName[0] - 'A' + 1) * 26 + (columnName[1] - 'A');
		}
		else
		{
			throw new ArgumentException("Invalid column name format");
		}
	}

	public void SetColumnWidth(int sheetId, int columnIndex, int width)
	{
		var updateDimensionPropertiesRequest = new Request
		{
			UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
			{
				Range = new DimensionRange
				{
					SheetId = sheetId,
					Dimension = "COLUMNS",
					StartIndex = columnIndex - 1,
					EndIndex = columnIndex
				},
				Properties = new DimensionProperties { PixelSize = width },
				Fields = "pixelSize"
			}
		};

		Requests.Add(updateDimensionPropertiesRequest);
	}

	public void BatchUpdate()
	{
		var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = Requests };

		bool success = false;
		int retries = 0;
		int maxRetries = 10; // Maximum number of retries

		while (!success && retries < maxRetries)
		{
			try
			{
				SheetService.Spreadsheets
					.BatchUpdate(batchUpdateRequest, SpreadSheetId)
					.Execute();

				success = true; // Request succeeded, exit loop
			}
			catch (Google.GoogleApiException ex)
			{
				if (ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests)
				{
					Console.WriteLine("Quota exceeded. Retrying in 30 seconds...");
					retries++;
					Thread.Sleep(30000); // Wait for 30 seconds before retrying
				}
				else
				{
					// Handle other exceptions
					Console.WriteLine("An error occurred: " + ex.Message);
					break;
				}
			}
			catch
			{
				// FUCK THIS SHIT
			}
		}

		Requests = new List<Request>();
	}

	public virtual void MarkSentEmailInExcel(int rowId)
	{
		var range = new ValueRange
		{
			Range = $"{SummaryTabName}!B{ROW_STARTING_INDEX + rowId}",
			MajorDimension = "COLUMNS",
			Values = new List<IList<object>> { new List<object> { $"{DateTime.Today.ToString("dd.MM.yyyy")}" } }
		};

		var updateRequest = SheetService.Spreadsheets.Values.Update(range, SelectedEvent?.SpreadsheetId, range.Range);
		updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

		var updateResponse = updateRequest.Execute();
	}

	public void ClearExcel()
	{
		var spreadsheet = SheetService.Spreadsheets.Get(SpreadSheetId).Execute();
		IList<Sheet> sheets = spreadsheet.Sheets;

		BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
		batchUpdateSpreadsheetRequest.Requests = new List<Request>();

		if (sheets.Count <= 1)
		{
			throw new Exception("Nie ma zakładek do usunięcia.");
		}

		foreach (var sheet in sheets)
		{
			if (sheet.Properties.Title != _rawDataTabName)
			{
				Request request = new Request()
				{
					DeleteSheet = new DeleteSheetRequest()
					{
						SheetId = sheet.Properties.SheetId
					}
				};

				batchUpdateSpreadsheetRequest.Requests.Add(request);
			}
		}

		SheetService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SpreadSheetId).Execute();
	}

	public string GetColumnName(int columnIndex)
	{
		const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		string columnLetter = string.Empty;

		while (columnIndex > 0)
		{
			columnIndex--; // Adjust for 0-based indexing
			int remainder = columnIndex % 26;
			columnLetter = alphabet[remainder] + columnLetter;
			columnIndex /= 26;
		}

		return columnLetter;
	}

	public int GetColumnIndex(string? columnLetter)
	{
		if (columnLetter == null)
			return 0;

		const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		columnLetter = columnLetter.ToUpper();
		int columnIndex = 0;

		for (int i = columnLetter.Length - 1; i >= 0; i--)
		{
			char currentChar = columnLetter[i];
			int positionInAlphabet = alphabet.IndexOf(currentChar) + 1;
			columnIndex += positionInAlphabet * (int)Math.Pow(26, columnLetter.Length - 1 - i);
		}

		return columnIndex;
	}

	public int? GetSheetIdForCurrentEvent(string sheetName)
		=> SheetService
			.Spreadsheets
			.Get(SelectedEvent?.SpreadsheetId)
			.Execute()
			.Sheets
			.SingleOrDefault(sheet => sheet.Properties.Title.Equals(sheetName))?
			.Properties
			.SheetId;

    public void GenerateTabs(Event selectedEvent)
    {
        SetUpRegistrationsEditableTab();
        SetUpSummaryTab();
        SetUpTabForAccountants();
        SetUpTabForNoPayments();
        SetUpGroupBalanceTab();
    }
}