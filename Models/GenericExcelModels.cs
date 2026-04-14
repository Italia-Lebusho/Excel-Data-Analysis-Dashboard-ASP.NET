namespace Markerting.Models
{
    public enum GenericColumnKind
    {
        Text,
        Number,
        Date
    }

    public enum AnalysisGoalKind
    {
        Overview = 0,
        CompareGroups = 1,
        BreakdownByCategory = 2,
        TrendOverTime = 3
    }

    /// <summary>User choices for guided analysis (stored in session after setup).</summary>
    public class UserAnalysisRequest
    {
        public AnalysisGoalKind Goal { get; set; } = AnalysisGoalKind.Overview;
        public string? CategoryColumn { get; set; }
        public string? MeasureColumn { get; set; }
        public string? DateColumn { get; set; }
        public string? CompareValueA { get; set; }
        public string? CompareValueB { get; set; }
    }

    public class AnalyzeSetupFormModel
    {
        public AnalysisGoalKind Goal { get; set; }
        public string? CategoryColumn { get; set; }
        public string? MeasureColumn { get; set; }
        public string? DateColumn { get; set; }
        public string? CompareValueA { get; set; }
        public string? CompareValueB { get; set; }
    }

    public class ColumnOptionVm
    {
        public string Name { get; set; } = string.Empty;
        public string Kind { get; set; } = string.Empty;
    }

    public class AnalyzeSetupViewModel
    {
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public List<ColumnOptionVm> Columns { get; set; } = new();
        public Dictionary<string, List<string>> Suggestions { get; set; } = new();
        public UserAnalysisRequest? Current { get; set; }
    }

    public class GenericComparisonRowVm
    {
        public string Metric { get; set; } = string.Empty;
        public string ValueA { get; set; } = string.Empty;
        public string ValueB { get; set; } = string.Empty;
        public string? Note { get; set; }
    }

    public class GenericColumnSummary
    {
        public string Name { get; set; } = string.Empty;
        public GenericColumnKind Kind { get; set; }
        public int NonEmptyCount { get; set; }
        public int DistinctApprox { get; set; }
        public decimal? Sum { get; set; }
        public decimal? Average { get; set; }
        public decimal? Min { get; set; }
        public decimal? Max { get; set; }
        public string? DateMin { get; set; }
        public string? DateMax { get; set; }
    }

    public class GenericChartDatasetVm
    {
        public string Label { get; set; } = string.Empty;
        public List<decimal> Data { get; set; } = new();
    }

    public class GenericChartVm
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string ChartType { get; set; } = "bar"; // bar | line | doughnut
        public List<string> Labels { get; set; } = new();
        public List<GenericChartDatasetVm> Datasets { get; set; } = new();
        public string? YAxisTickPrefix { get; set; }
    }

    public class GenericKpiVm
    {
        public string Label { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string CssClass { get; set; } = "bg-primary";
    }

    /// <summary>Short narrative + practical next steps generated from the sheet profile.</summary>
    public class GenericInsightStoryVm
    {
        public string Teaser { get; set; } = string.Empty;
        public List<string> Observations { get; set; } = new();
        public List<string> Recommendations { get; set; } = new();
    }

    public class GenericDashboardViewModel
    {
        public bool HasData { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public int TruncatedRowCount { get; set; }
        public List<GenericKpiVm> Kpis { get; set; } = new();
        public List<GenericColumnSummary> Columns { get; set; } = new();
        public List<GenericChartVm> Charts { get; set; } = new();
        public List<Dictionary<string, string>> PreviewRows { get; set; } = new();
        public GenericInsightStoryVm? InsightStory { get; set; }
        public UserAnalysisRequest? AnalysisRequest { get; set; }
        public string? AnalysisPlanSummary { get; set; }
        public List<GenericComparisonRowVm> FinalComparison { get; set; } = new();
    }

    /// <summary>Serializable payload stored in session for re-analysis on each request.</summary>
    public class GenericExcelSessionPayload
    {
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public List<string> Headers { get; set; } = new();
        public List<List<string>> Rows { get; set; } = new();
        public bool WasTruncated { get; set; }
    }
}
