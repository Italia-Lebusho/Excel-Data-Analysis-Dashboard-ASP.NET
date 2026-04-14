using System.Globalization;
using System.Text.Json;
using Markerting.Models;
using OfficeOpenXml;

namespace Markerting.Services
{
    public class GenericExcelAnalysisService
    {
        public const int MaxRowsStored = 4000;

        public GenericExcelSessionPayload? ParseUpload(Stream stream, string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];
            var dim = worksheet.Dimension;
            if (dim == null || dim.Rows < 2)
                return null;

            var headerRow = DetectHeaderRow(worksheet, dim.Rows, dim.Columns);
            var colCount = dim.Columns;
            var headers = new List<string>();
            for (int c = 1; c <= colCount; c++)
            {
                var h = worksheet.Cells[headerRow, c].Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(h))
                    h = $"Column{c}";
                headers.Add(h);
            }
            DeduplicateHeaders(headers);

            var rows = new List<List<string>>();
            var totalBodyRows = 0;
            for (int r = headerRow + 1; r <= dim.Rows; r++)
            {
                if (IsRowEffectivelyEmpty(worksheet, r, colCount))
                    continue;
                totalBodyRows++;
                if (rows.Count >= MaxRowsStored)
                    continue;
                var row = new List<string>(colCount);
                for (int c = 1; c <= colCount; c++)
                    row.Add(worksheet.Cells[r, c].Text?.Trim() ?? "");
                rows.Add(row);
            }

            return new GenericExcelSessionPayload
            {
                FileName = fileName,
                SheetName = worksheet.Name,
                Headers = headers,
                Rows = rows,
                WasTruncated = totalBodyRows > rows.Count
            };
        }

        public GenericDashboardViewModel BuildDashboard(GenericExcelSessionPayload payload, UserAnalysisRequest? request = null)
        {
            var vm = new GenericDashboardViewModel
            {
                HasData = true,
                FileName = payload.FileName,
                SheetName = payload.SheetName,
                RowCount = payload.Rows.Count,
                ColumnCount = payload.Headers.Count,
                TruncatedRowCount = 0,
                AnalysisRequest = request,
                AnalysisPlanSummary = DescribeAnalysisPlan(request)
            };

            if (payload.Headers.Count == 0 || payload.Rows.Count == 0)
            {
                vm.HasData = false;
                return vm;
            }

            var colKinds = InferColumnKinds(payload);
            var numericValues = new List<decimal?>[payload.Headers.Count];
            var dateValues = new List<DateTime?>[payload.Headers.Count];
            for (int i = 0; i < payload.Headers.Count; i++)
            {
                numericValues[i] = new List<decimal?>(payload.Rows.Count);
                dateValues[i] = new List<DateTime?>(payload.Rows.Count);
            }

            foreach (var row in payload.Rows)
            {
                for (int c = 0; c < payload.Headers.Count; c++)
                {
                    var cell = c < row.Count ? row[c] : "";
                    numericValues[c].Add(TryParseDecimal(cell, out var dec) ? dec : null);
                    dateValues[c].Add(TryParseDate(cell, out var dt) ? dt : null);
                }
            }

            for (int c = 0; c < payload.Headers.Count; c++)
            {
                var summary = new GenericColumnSummary { Name = payload.Headers[c], Kind = colKinds[c] };
                var nums = numericValues[c].Where(x => x.HasValue).Select(x => x!.Value).ToList();
                var dates = dateValues[c].Where(x => x.HasValue).Select(x => x!.Value).ToList();
                summary.NonEmptyCount = payload.Rows.Count(r => c < r.Count && !string.IsNullOrWhiteSpace(r[c]));

                if (colKinds[c] == GenericColumnKind.Number && nums.Count > 0)
                {
                    summary.Sum = nums.Sum();
                    summary.Average = nums.Average();
                    summary.Min = nums.Min();
                    summary.Max = nums.Max();
                    summary.DistinctApprox = nums.Select(x => x.ToString(CultureInfo.InvariantCulture)).Distinct().Count();
                }
                else if (colKinds[c] == GenericColumnKind.Date && dates.Count > 0)
                {
                    summary.DateMin = dates.Min().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    summary.DateMax = dates.Max().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    summary.DistinctApprox = dates.Select(d => d.Date).Distinct().Count();
                }
                else
                {
                    var texts = payload.Rows
                        .Select(r => c < r.Count ? r[c]?.Trim() ?? "" : "")
                        .Where(t => t.Length > 0)
                        .ToList();
                    summary.DistinctApprox = texts.Distinct(StringComparer.OrdinalIgnoreCase).Count();
                }

                vm.Columns.Add(summary);
            }

            vm.Kpis.Add(new GenericKpiVm { Label = "Rows analyzed", Value = vm.RowCount.ToString("N0"), CssClass = "bg-primary" });
            vm.Kpis.Add(new GenericKpiVm { Label = "Columns", Value = vm.ColumnCount.ToString(), CssClass = "bg-secondary" });

            var chartSeq = 0;
            void AddChart(GenericChartVm chart)
            {
                chart.Id = $"chart_{chartSeq++}";
                vm.Charts.Add(chart);
            }

            if (request == null || request.Goal == AnalysisGoalKind.Overview)
            {
                foreach (var col in vm.Columns.Where(x => x.Kind == GenericColumnKind.Number && x.Sum.HasValue))
                    vm.Kpis.Add(new GenericKpiVm { Label = $"Sum · {col.Name}", Value = FormatNumber(col.Sum!.Value), CssClass = "bg-success" });
                AddDefaultDiscoveryCharts(vm, payload, colKinds, AddChart);
            }
            else
            {
                ApplyGuidedAnalysis(vm, payload, colKinds, request, AddChart);
            }

            FillPreviewRows(vm, payload, request);
            vm.InsightStory = BuildInsightStory(vm, payload, colKinds, request);
            return vm;
        }

        private static string? DescribeAnalysisPlan(UserAnalysisRequest? request)
        {
            if (request == null || request.Goal == AnalysisGoalKind.Overview)
                return "Open overview: scan every column type and surface the strongest automatic charts.";

            return request.Goal switch
            {
                AnalysisGoalKind.CompareGroups =>
                    $"Compare “{request.CompareValueA}” vs “{request.CompareValueB}” within “{request.CategoryColumn}”" +
                    (string.IsNullOrWhiteSpace(request.MeasureColumn) ? " using row counts." : $" using totals from “{request.MeasureColumn}”."),
                AnalysisGoalKind.BreakdownByCategory =>
                    $"Rank categories in “{request.CategoryColumn}”" +
                    (string.IsNullOrWhiteSpace(request.MeasureColumn) ? " by how often they appear." : $" by the sum of “{request.MeasureColumn}”."),
                AnalysisGoalKind.TrendOverTime =>
                    $"Show the time trend for “{request.DateColumn}”" +
                    (string.IsNullOrWhiteSpace(request.MeasureColumn) ? " (row counts per month)." : $" with totals from “{request.MeasureColumn}” each month."),
                _ => null
            };
        }

        private static void AddDefaultDiscoveryCharts(
            GenericDashboardViewModel vm,
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            Action<GenericChartVm> addChart)
        {
            static bool IsAverageOnlyColumn(string header) =>
                string.Equals(header, "Team", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(header, "AgentID", StringComparison.OrdinalIgnoreCase);

            int? dateColIndex = null;
            for (int i = 0; i < vm.Columns.Count; i++)
            {
                if (vm.Columns[i].Kind == GenericColumnKind.Date)
                {
                    dateColIndex = i;
                    break;
                }
            }

            if (dateColIndex is int dateCol)
            {
                var byMonth = payload.Rows
                    .Select(r => dateCol < r.Count ? r[dateCol] : "")
                    .Select(cell => TryParseDate(cell, out var dt) ? dt : (DateTime?)null)
                    .Where(dt => dt.HasValue)
                    .GroupBy(dt => new DateTime(dt!.Value.Year, dt.Value.Month, 1))
                    .OrderBy(g => g.Key)
                    .Select(g => new
                    {
                        Label = g.Key.ToString("MMM yyyy", CultureInfo.InvariantCulture),
                        Count = g.Count()
                    })
                    .ToList();
                if (byMonth.Count > 0)
                {
                    addChart(new GenericChartVm
                    {
                        Title = $"Activity over time · {payload.Headers[dateCol]}",
                        ChartType = "line",
                        Labels = byMonth.Select(x => x.Label).ToList(),
                        Datasets = new List<GenericChartDatasetVm>
                        {
                            new()
                            {
                                Label = "Row count",
                                Data = byMonth.Select(x => (decimal)x.Count).ToList()
                            }
                        }
                    });
                }
            }

            int? measureColIndex = null;
            for (int i = 0; i < vm.Columns.Count; i++)
            {
                if (vm.Columns[i].Kind == GenericColumnKind.Number)
                {
                    measureColIndex = i;
                    break;
                }
            }

            var categoryCandidates = vm.Columns
                .Select((c, i) => (c, i))
                .Where(x => x.c.Kind == GenericColumnKind.Text && x.c.DistinctApprox >= 2 && x.c.DistinctApprox <= 80)
                .Select(x => x.i)
                .ToList();

            foreach (var catCol in categoryCandidates.Take(3))
            {
                if (vm.Charts.Count >= 8) break;
                if (IsAverageOnlyColumn(payload.Headers[catCol])) continue;
                var counts = payload.Rows
                    .Select(r => catCol < r.Count ? r[catCol]?.Trim() ?? "" : "")
                    .Where(s => s.Length > 0)
                    .GroupBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(g => g.Count())
                    .Take(14)
                    .ToList();
                if (counts.Count == 0) continue;
                addChart(new GenericChartVm
                {
                    Title = $"Top values · {payload.Headers[catCol]}",
                    ChartType = counts.Count <= 6 ? "doughnut" : "bar",
                    Labels = counts.Select(g => g.Key.Length > 40 ? g.Key[..37] + "…" : g.Key).ToList(),
                    Datasets = new List<GenericChartDatasetVm>
                    {
                        new() { Label = "Count", Data = counts.Select(g => (decimal)g.Count()).ToList() }
                    }
                });
            }

            if (measureColIndex is int measureCol && categoryCandidates.Count > 0)
            {
                var rankedCats = categoryCandidates
                    .Where(i => i != measureCol)
                    .OrderBy(i => vm.Columns[i].DistinctApprox)
                    .ToList();
                if (rankedCats.Count > 0)
                {
                    var catIdx = rankedCats[0];
                    var sums = payload.Rows
                        .GroupBy(r =>
                        {
                            var key = catIdx < r.Count ? r[catIdx]?.Trim() ?? "" : "";
                            return string.IsNullOrEmpty(key) ? "(blank)" : key;
                        }, StringComparer.OrdinalIgnoreCase)
                        .Select(g => new
                        {
                            Key = g.Key.Length > 36 ? g.Key[..33] + "…" : g.Key,
                            Sum = g.Sum(row => measureCol < row.Count && TryParseDecimal(row[measureCol], out var v) ? v : 0m)
                        })
                        .OrderByDescending(x => x.Sum)
                        .Take(12)
                        .ToList();
                    if (sums.Count > 0 && sums.Any(x => x.Sum != 0))
                    {
                        addChart(new GenericChartVm
                        {
                            Title = $"Sum of {payload.Headers[measureCol]} by {payload.Headers[catIdx]}",
                            ChartType = "bar",
                            Labels = sums.Select(x => x.Key).ToList(),
                            Datasets = new List<GenericChartDatasetVm>
                            {
                                new() { Label = payload.Headers[measureCol], Data = sums.Select(x => x.Sum).ToList() }
                            },
                            YAxisTickPrefix = ""
                        });
                    }
                }

                // For Team and AgentID, show averages instead of top-value counts.
                foreach (var catIdx in categoryCandidates.Where(i => IsAverageOnlyColumn(payload.Headers[i])).Take(2))
                {
                    if (vm.Charts.Count >= 8) break;
                    var averages = payload.Rows
                        .GroupBy(r =>
                        {
                            var key = catIdx < r.Count ? r[catIdx]?.Trim() ?? "" : "";
                            return string.IsNullOrEmpty(key) ? "(blank)" : key;
                        }, StringComparer.OrdinalIgnoreCase)
                        .Select(g =>
                        {
                            var values = g
                                .Select(row => measureCol < row.Count && TryParseDecimal(row[measureCol], out var v) ? (decimal?)v : null)
                                .Where(v => v.HasValue)
                                .Select(v => v!.Value)
                                .ToList();
                            return new
                            {
                                Key = g.Key.Length > 36 ? g.Key[..33] + "…" : g.Key,
                                Avg = values.Count > 0 ? values.Average() : 0m,
                                Count = values.Count
                            };
                        })
                        .Where(x => x.Count > 0)
                        .OrderByDescending(x => x.Avg)
                        .Take(12)
                        .ToList();

                    if (averages.Count == 0) continue;

                    addChart(new GenericChartVm
                    {
                        Title = $"Average {payload.Headers[measureCol]} by {payload.Headers[catIdx]}",
                        ChartType = "bar",
                        Labels = averages.Select(x => x.Key).ToList(),
                        Datasets = new List<GenericChartDatasetVm>
                        {
                            new() { Label = $"Avg {payload.Headers[measureCol]}", Data = averages.Select(x => x.Avg).ToList() }
                        }
                    });
                }
            }

            if (vm.Columns.Count(x => x.Kind == GenericColumnKind.Number) > 1)
            {
                var totals = vm.Columns
                    .Where(x => x.Kind == GenericColumnKind.Number && x.Sum.HasValue)
                    .Take(10)
                    .Select(x => (x.Name, x.Sum!.Value))
                    .ToList();
                if (totals.Count > 0)
                {
                    addChart(new GenericChartVm
                    {
                        Title = "Column totals (numeric)",
                        ChartType = "bar",
                        Labels = totals.Select(t => t.Name.Length > 28 ? t.Name[..25] + "…" : t.Name).ToList(),
                        Datasets = new List<GenericChartDatasetVm>
                        {
                            new() { Label = "Sum", Data = totals.Select(t => t.Value).ToList() }
                        }
                    });
                }
            }
        }

        private void ApplyGuidedAnalysis(
            GenericDashboardViewModel vm,
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            UserAnalysisRequest request,
            Action<GenericChartVm> addChart)
        {
            switch (request.Goal)
            {
                case AnalysisGoalKind.CompareGroups:
                    ApplyCompareGroups(vm, payload, colKinds, request, addChart);
                    break;
                case AnalysisGoalKind.BreakdownByCategory:
                    ApplyBreakdown(vm, payload, colKinds, request, addChart);
                    break;
                case AnalysisGoalKind.TrendOverTime:
                    ApplyTrend(vm, payload, colKinds, request, addChart);
                    break;
            }
        }

        private void ApplyCompareGroups(
            GenericDashboardViewModel vm,
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            UserAnalysisRequest request,
            Action<GenericChartVm> addChart)
        {
            var catIdx = FindColumnIndex(payload.Headers, request.CategoryColumn);
            if (catIdx < 0 || colKinds[catIdx] != GenericColumnKind.Text) return;

            var a = request.CompareValueA?.Trim() ?? "";
            var b = request.CompareValueB?.Trim() ?? "";
            int? measureIdx = string.IsNullOrWhiteSpace(request.MeasureColumn)
                ? null
                : FindColumnIndex(payload.Headers, request.MeasureColumn);
            if (measureIdx is int mi && (mi < 0 || colKinds[mi] != GenericColumnKind.Number))
                measureIdx = null;

            bool Match(List<string> row, string value) =>
                string.Equals(Cell(row, catIdx), value, StringComparison.OrdinalIgnoreCase);

            var rowsA = payload.Rows.Where(r => Match(r, a)).ToList();
            var rowsB = payload.Rows.Where(r => Match(r, b)).ToList();
            var ca = rowsA.Count;
            var cb = rowsB.Count;

            decimal SumMeasure(List<List<string>> rows) =>
                rows.Sum(r => measureIdx is int ix && ix < r.Count && TryParseDecimal(r[ix], out var v) ? v : 0m);

            var sumA = measureIdx.HasValue ? SumMeasure(rowsA) : 0m;
            var sumB = measureIdx.HasValue ? SumMeasure(rowsB) : 0m;
            var avgA = measureIdx.HasValue && ca > 0 ? sumA / ca : 0m;
            var avgB = measureIdx.HasValue && cb > 0 ? sumB / cb : 0m;

            vm.Kpis.Add(new GenericKpiVm { Label = $"Rows · {TruncateLabel(a, 24)}", Value = ca.ToString("N0"), CssClass = "bg-info" });
            vm.Kpis.Add(new GenericKpiVm { Label = $"Rows · {TruncateLabel(b, 24)}", Value = cb.ToString("N0"), CssClass = "bg-warning text-dark" });
            if (measureIdx.HasValue)
            {
                vm.Kpis.Add(new GenericKpiVm { Label = $"Total · {payload.Headers[measureIdx.Value]} · A", Value = FormatNumber(sumA), CssClass = "bg-success" });
                vm.Kpis.Add(new GenericKpiVm { Label = $"Total · {payload.Headers[measureIdx.Value]} · B", Value = FormatNumber(sumB), CssClass = "bg-secondary" });
            }

            if (measureIdx.HasValue)
            {
                addChart(new GenericChartVm
                {
                    Title = $"Side-by-side · {payload.Headers[measureIdx.Value]}",
                    ChartType = "bar",
                    Labels = new List<string> { "Rows matched", $"Sum · {payload.Headers[measureIdx.Value]}" },
                    Datasets = new List<GenericChartDatasetVm>
                    {
                        new() { Label = TruncateLabel(a, 20), Data = new List<decimal> { ca, sumA } },
                        new() { Label = TruncateLabel(b, 20), Data = new List<decimal> { cb, sumB } }
                    }
                });
            }
            else
            {
                addChart(new GenericChartVm
                {
                    Title = "Side-by-side row counts",
                    ChartType = "bar",
                    Labels = new List<string> { TruncateLabel(a, 28), TruncateLabel(b, 28) },
                    Datasets = new List<GenericChartDatasetVm>
                    {
                        new() { Label = "Rows", Data = new List<decimal> { ca, cb } }
                    }
                });
            }

            vm.FinalComparison.Add(new GenericComparisonRowVm { Metric = "Rows matched", ValueA = ca.ToString("N0"), ValueB = cb.ToString("N0"), Note = ca == 0 || cb == 0 ? "One side had zero matches—check spelling or extra spaces." : null });
            if (measureIdx.HasValue)
            {
                vm.FinalComparison.Add(new GenericComparisonRowVm { Metric = $"Total · {payload.Headers[measureIdx.Value]}", ValueA = FormatNumber(sumA), ValueB = FormatNumber(sumB) });
                vm.FinalComparison.Add(new GenericComparisonRowVm { Metric = $"Average · {payload.Headers[measureIdx.Value]}", ValueA = ca > 0 ? FormatNumber(avgA) : "—", ValueB = cb > 0 ? FormatNumber(avgB) : "—" });
            }
        }

        private void ApplyBreakdown(
            GenericDashboardViewModel vm,
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            UserAnalysisRequest request,
            Action<GenericChartVm> addChart)
        {
            var catIdx = FindColumnIndex(payload.Headers, request.CategoryColumn);
            if (catIdx < 0 || colKinds[catIdx] != GenericColumnKind.Text) return;

            int? measureIdx = string.IsNullOrWhiteSpace(request.MeasureColumn)
                ? null
                : FindColumnIndex(payload.Headers, request.MeasureColumn);
            if (measureIdx is int mi && (mi < 0 || colKinds[mi] != GenericColumnKind.Number))
                measureIdx = null;

            var groups = payload.Rows
                .GroupBy(r => string.IsNullOrWhiteSpace(Cell(r, catIdx)) ? "(blank)" : Cell(r, catIdx)!, StringComparer.OrdinalIgnoreCase)
                .Select(g => new
                {
                    Key = g.Key.Length > 36 ? g.Key[..33] + "…" : g.Key,
                    Total = measureIdx.HasValue
                        ? g.Sum(row => measureIdx.Value < row.Count && TryParseDecimal(row[measureIdx.Value], out var v) ? v : 0m)
                        : (decimal)g.Count()
                })
                .OrderByDescending(x => x.Total)
                .ToList();
            if (groups.Count == 0) return;

            var grand = groups.Sum(x => x.Total);
            var top = groups.Take(16).ToList();
            addChart(new GenericChartVm
            {
                Title = measureIdx.HasValue
                    ? $"Totals by {payload.Headers[catIdx]} · {payload.Headers[measureIdx.Value]}"
                    : $"Frequency by {payload.Headers[catIdx]}",
                ChartType = "bar",
                Labels = top.Select(x => x.Key).ToList(),
                Datasets = new List<GenericChartDatasetVm>
                {
                    new()
                    {
                        Label = measureIdx.HasValue ? payload.Headers[measureIdx.Value] : "Rows",
                        Data = top.Select(x => (decimal)x.Total).ToList()
                    }
                }
            });

            vm.Kpis.Add(new GenericKpiVm { Label = "Categories", Value = groups.Count.ToString("N0"), CssClass = "bg-info" });
            vm.Kpis.Add(new GenericKpiVm { Label = "Top category", Value = TruncateLabel(top[0].Key, 32), CssClass = "bg-success" });

            foreach (var row in top.Take(8))
            {
                var share = grand > 0 ? (row.Total / grand * 100m) : 0m;
                vm.FinalComparison.Add(new GenericComparisonRowVm
                {
                    Metric = row.Key,
                    ValueA = measureIdx.HasValue ? FormatNumber(row.Total) : row.Total.ToString("N0"),
                    ValueB = $"{share:F1}% of total"
                });
            }
        }

        private void ApplyTrend(
            GenericDashboardViewModel vm,
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            UserAnalysisRequest request,
            Action<GenericChartVm> addChart)
        {
            var dateIdx = FindColumnIndex(payload.Headers, request.DateColumn);
            if (dateIdx < 0 || colKinds[dateIdx] != GenericColumnKind.Date) return;

            int? measureIdx = string.IsNullOrWhiteSpace(request.MeasureColumn)
                ? null
                : FindColumnIndex(payload.Headers, request.MeasureColumn);
            if (measureIdx is int mx && (mx < 0 || colKinds[mx] != GenericColumnKind.Number))
                measureIdx = null;

            var buckets = new Dictionary<DateTime, (decimal Value, int Rows)>();
            foreach (var row in payload.Rows)
            {
                var cell = dateIdx < row.Count ? row[dateIdx] : "";
                if (!TryParseDate(cell, out var dt)) continue;
                var key = new DateTime(dt.Year, dt.Month, 1);
                var cur = buckets.GetValueOrDefault(key);
                if (measureIdx.HasValue)
                {
                    var add = measureIdx.Value < row.Count && TryParseDecimal(row[measureIdx.Value], out var v) ? v : 0m;
                    buckets[key] = (cur.Value + add, cur.Rows + 1);
                }
                else
                {
                    buckets[key] = (cur.Value + 1m, cur.Rows + 1);
                }
            }
            if (buckets.Count == 0) return;

            var ordered = buckets.OrderBy(kv => kv.Key).ToList();
            var labels = ordered.Select(kv => kv.Key.ToString("MMM yyyy", CultureInfo.InvariantCulture)).ToList();
            var values = ordered.Select(kv => kv.Value.Value).ToList();

            addChart(new GenericChartVm
            {
                Title = measureIdx.HasValue
                    ? $"Trend · {payload.Headers[measureIdx.Value]} by month"
                    : $"Row activity by month · {payload.Headers[dateIdx]}",
                ChartType = "line",
                Labels = labels,
                Datasets = new List<GenericChartDatasetVm>
                {
                    new()
                    {
                        Label = measureIdx.HasValue ? payload.Headers[measureIdx.Value] : "Rows",
                        Data = values
                    }
                }
            });

            var first = values[0];
            var last = values[^1];
            vm.Kpis.Add(new GenericKpiVm { Label = "Months in view", Value = ordered.Count.ToString("N0"), CssClass = "bg-info" });
            vm.Kpis.Add(new GenericKpiVm { Label = "First month", Value = FormatNumber(first), CssClass = "bg-secondary" });
            vm.Kpis.Add(new GenericKpiVm { Label = "Last month", Value = FormatNumber(last), CssClass = "bg-success" });

            var delta = last - first;
            vm.FinalComparison.Add(new GenericComparisonRowVm { Metric = "First month in chart", ValueA = labels[0], ValueB = FormatNumber(first) });
            vm.FinalComparison.Add(new GenericComparisonRowVm { Metric = "Last month in chart", ValueA = labels[^1], ValueB = FormatNumber(last) });
            vm.FinalComparison.Add(new GenericComparisonRowVm { Metric = "Change (last − first)", ValueA = FormatNumber(delta), ValueB = first == 0 ? "—" : $"{(delta / Math.Abs(first) * 100m):F1}% vs start" });

            foreach (var kv in ordered.Take(6))
                vm.FinalComparison.Add(new GenericComparisonRowVm
                {
                    Metric = kv.Key.ToString("MMM yyyy", CultureInfo.InvariantCulture),
                    ValueA = measureIdx.HasValue ? FormatNumber(kv.Value.Value) : kv.Value.Rows.ToString("N0"),
                    ValueB = "Month total"
                });
        }

        private static void FillPreviewRows(GenericDashboardViewModel vm, GenericExcelSessionPayload payload, UserAnalysisRequest? request)
        {
            IEnumerable<List<string>> rows = payload.Rows;
            if (request?.Goal == AnalysisGoalKind.CompareGroups)
            {
                var catIdx = FindColumnIndex(payload.Headers, request.CategoryColumn);
                var a = request.CompareValueA?.Trim() ?? "";
                var b = request.CompareValueB?.Trim() ?? "";
                if (catIdx >= 0 && (a.Length > 0 || b.Length > 0))
                {
                    rows = payload.Rows.Where(r =>
                        string.Equals(Cell(r, catIdx), a, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(Cell(r, catIdx), b, StringComparison.OrdinalIgnoreCase));
                }
            }

            foreach (var row in rows.Take(25))
            {
                var dict = new Dictionary<string, string>(StringComparer.Ordinal);
                for (int c = 0; c < payload.Headers.Count; c++)
                {
                    var v = c < row.Count ? row[c] : "";
                    dict[payload.Headers[c]] = v.Length > 120 ? v[..117] + "…" : v;
                }
                vm.PreviewRows.Add(dict);
            }
        }

        private static int FindColumnIndex(IReadOnlyList<string> headers, string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return -1;
            for (var i = 0; i < headers.Count; i++)
            {
                if (string.Equals(headers[i], name.Trim(), StringComparison.Ordinal))
                    return i;
            }
            return -1;
        }

        private static string Cell(List<string> row, int col) =>
            col < row.Count ? row[col]?.Trim() ?? "" : "";

        public static string SerializeAnalysisRequest(UserAnalysisRequest request) =>
            JsonSerializer.Serialize(request);

        public static UserAnalysisRequest? DeserializeAnalysisRequest(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try
            {
                return JsonSerializer.Deserialize<UserAnalysisRequest>(json);
            }
            catch
            {
                return null;
            }
        }

        public AnalyzeSetupViewModel BuildAnalyzeSetupViewModel(GenericExcelSessionPayload payload, UserAnalysisRequest? current)
        {
            var kinds = InferColumnKinds(payload);
            var vm = new AnalyzeSetupViewModel
            {
                FileName = payload.FileName,
                SheetName = payload.SheetName,
                RowCount = payload.Rows.Count,
                Current = current
            };
            for (var i = 0; i < payload.Headers.Count; i++)
            {
                vm.Columns.Add(new ColumnOptionVm
                {
                    Name = payload.Headers[i],
                    Kind = kinds[i].ToString()
                });
            }

            for (var c = 0; c < payload.Headers.Count; c++)
            {
                if (kinds[c] != GenericColumnKind.Text) continue;
                var distinct = CollectDistinctValues(payload, c, 80);
                if (distinct.Count > 0)
                    vm.Suggestions[payload.Headers[c]] = distinct;
            }

            return vm;
        }

        private static List<string> CollectDistinctValues(GenericExcelSessionPayload payload, int colIndex, int max)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in payload.Rows)
            {
                var v = colIndex < row.Count ? row[colIndex]?.Trim() ?? "" : "";
                if (v.Length == 0) continue;
                set.Add(v);
                if (set.Count >= max) break;
            }
            return set.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).Take(max).ToList();
        }

        public string? ValidateAnalysisForm(GenericExcelSessionPayload payload, AnalyzeSetupFormModel form, out UserAnalysisRequest request)
        {
            request = new UserAnalysisRequest
            {
                Goal = form.Goal,
                CategoryColumn = string.IsNullOrWhiteSpace(form.CategoryColumn) ? null : form.CategoryColumn.Trim(),
                MeasureColumn = string.IsNullOrWhiteSpace(form.MeasureColumn) ? null : form.MeasureColumn.Trim(),
                DateColumn = string.IsNullOrWhiteSpace(form.DateColumn) ? null : form.DateColumn.Trim(),
                CompareValueA = string.IsNullOrWhiteSpace(form.CompareValueA) ? null : form.CompareValueA.Trim(),
                CompareValueB = string.IsNullOrWhiteSpace(form.CompareValueB) ? null : form.CompareValueB.Trim()
            };

            if (request.Goal == AnalysisGoalKind.Overview)
                return null;

            var kinds = InferColumnKinds(payload);

            if (request.Goal == AnalysisGoalKind.CompareGroups)
            {
                if (string.IsNullOrEmpty(request.CategoryColumn))
                    return "Choose the text column you want to split on.";
                var ci = FindColumnIndex(payload.Headers, request.CategoryColumn);
                if (ci < 0) return $"Could not find a column named “{request.CategoryColumn}”.";
                if (kinds[ci] != GenericColumnKind.Text)
                    return "Compare two groups only works when the split column is treated as text (re-type mixed columns if needed).";
                if (string.IsNullOrEmpty(request.CompareValueA) || string.IsNullOrEmpty(request.CompareValueB))
                    return "Enter both values to compare (exact text, case-insensitive match).";
                if (string.Equals(request.CompareValueA, request.CompareValueB, StringComparison.OrdinalIgnoreCase))
                    return "Pick two different values to compare.";
                if (!string.IsNullOrEmpty(request.MeasureColumn))
                {
                    var mi = FindColumnIndex(payload.Headers, request.MeasureColumn);
                    if (mi < 0) return $"Could not find measure column “{request.MeasureColumn}”.";
                    if (kinds[mi] != GenericColumnKind.Number)
                        return "The measure column must be numeric.";
                }
                if (!string.IsNullOrEmpty(request.DateColumn))
                    return "Date column is not used for compare mode—clear it or switch to Trend over time.";
                return null;
            }

            if (request.Goal == AnalysisGoalKind.BreakdownByCategory)
            {
                if (string.IsNullOrEmpty(request.CategoryColumn))
                    return "Choose the category column.";
                var ci = FindColumnIndex(payload.Headers, request.CategoryColumn);
                if (ci < 0) return $"Could not find column “{request.CategoryColumn}”.";
                if (kinds[ci] != GenericColumnKind.Text)
                    return "Breakdown expects a text category column.";
                if (!string.IsNullOrEmpty(request.MeasureColumn))
                {
                    var mi = FindColumnIndex(payload.Headers, request.MeasureColumn);
                    if (mi < 0) return $"Could not find measure column “{request.MeasureColumn}”.";
                    if (kinds[mi] != GenericColumnKind.Number)
                        return "Measure must be a numeric column.";
                }
                if (!string.IsNullOrEmpty(request.DateColumn))
                    return "Date column is not used for breakdown—clear it or pick Trend over time.";
                if (!string.IsNullOrEmpty(request.CompareValueA) || !string.IsNullOrEmpty(request.CompareValueB))
                    return "Compare values are only used in Compare two groups—clear them for breakdown.";
                return null;
            }

            if (request.Goal == AnalysisGoalKind.TrendOverTime)
            {
                if (string.IsNullOrEmpty(request.DateColumn))
                    return "Pick the date column for the timeline.";
                var di = FindColumnIndex(payload.Headers, request.DateColumn);
                if (di < 0) return $"Could not find column “{request.DateColumn}”.";
                if (kinds[di] != GenericColumnKind.Date)
                    return "Trend mode needs a column detected as dates. Re-save dates as real date cells in Excel if this stays wrong.";
                if (!string.IsNullOrEmpty(request.MeasureColumn))
                {
                    var mi = FindColumnIndex(payload.Headers, request.MeasureColumn);
                    if (mi < 0) return $"Could not find measure column “{request.MeasureColumn}”.";
                    if (kinds[mi] != GenericColumnKind.Number)
                        return "Measure must be numeric for a value trend.";
                }
                if (!string.IsNullOrEmpty(request.CategoryColumn))
                    return "Category column is not used for trend—clear it or choose another mode.";
                if (!string.IsNullOrEmpty(request.CompareValueA) || !string.IsNullOrEmpty(request.CompareValueB))
                    return "Compare values are not used for trend—clear them.";
                return null;
            }

            return "Unknown analysis goal.";
        }

        private GenericInsightStoryVm BuildInsightStory(
            GenericDashboardViewModel vm,
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            UserAnalysisRequest? request)
        {
            var n = Math.Max(vm.RowCount, 1);
            var observations = new List<string>();
            var tips = new List<string>();

            if (request != null && request.Goal != AnalysisGoalKind.Overview && !string.IsNullOrWhiteSpace(vm.AnalysisPlanSummary))
                observations.Add("Your stated focus: " + vm.AnalysisPlanSummary);

            observations.Add(
                $"You are looking at {vm.RowCount:N0} rows and {vm.ColumnCount} columns from “{vm.SheetName}”.{(payload.WasTruncated ? " Very large files are sampled for speed, so re-run analysis in a dedicated BI tool if you need the full tail of the data." : string.Empty)}");

            if (request?.Goal == AnalysisGoalKind.CompareGroups)
                tips.Add("If the split looks wrong, return to “Choose what to analyze” and match the exact spelling or casing of the two values in your category column.");
            if (request?.Goal == AnalysisGoalKind.BreakdownByCategory)
                tips.Add("Take the top three slices from the comparison table and give each a clear owner and deadline.");
            if (request?.Goal == AnalysisGoalKind.TrendOverTime)
                tips.Add("Overlay major events (campaigns, outages, price changes) on the same months to explain spikes or dips.");

            var worstMissing = vm.Columns
                .Select(c => (c, Pct: (double)(n - c.NonEmptyCount) / n))
                .OrderByDescending(x => x.Pct)
                .FirstOrDefault();
            if (worstMissing.c != null && worstMissing.Pct >= 0.3)
            {
                observations.Add(
                    $"“{worstMissing.c.Name}” is empty on about {worstMissing.Pct * 100:F0}% of rows. Gaps like that quietly shrink the value of filters, sums, and “top category” charts.");
                tips.Add(
                    $"Solve it at the source: make “{worstMissing.c.Name}” required where it should never be blank, or agree a single default (for example “Unknown”) so every row stays comparable.");
            }

            var topNumeric = vm.Columns
                .Where(c => c.Kind == GenericColumnKind.Number && c.Sum.HasValue)
                .OrderByDescending(c => Math.Abs(c.Sum!.Value))
                .FirstOrDefault();
            if (topNumeric != null)
                observations.Add(
                    $"The strongest numeric signal is “{topNumeric.Name}” with a row-level total near {FormatNumber(topNumeric.Sum!.Value)}. That is the column to anchor targets, budgets, or exception alerts to.");

            foreach (var col in vm.Columns.Where(c => c.Kind == GenericColumnKind.Number && c.Average.HasValue && c.Max.HasValue && c.Min.HasValue && Math.Abs(c.Average!.Value) > 0.0001m))
            {
                var span = col.Max!.Value - col.Min!.Value;
                var ratio = Math.Abs(span / col.Average!.Value);
                if (ratio < 12m) continue;
                observations.Add(
                    $"“{col.Name}” swings widely relative to its typical value (min {FormatNumber(col.Min!.Value)}, max {FormatNumber(col.Max!.Value)}). A few extreme rows may be driving the story.");
                tips.Add(
                    $"Sort by “{col.Name}”, review the top and bottom 1–2%, and confirm whether those are errors, one-off deals, or real outliers you should model separately.");
                break;
            }

            var (skewCol, skewLabel, skewShare) = FindDominantTextCategory(payload, colKinds, vm.Columns);
            if (skewCol != null && skewShare >= 0.5)
            {
                observations.Add(
                    $"In “{skewCol}”, the value “{TruncateLabel(skewLabel, 48)}” appears on about {skewShare * 100:F0}% of filled rows. That is either a genuine leader or a sign that the field is under-used.");
                tips.Add(
                    $"If concentration is unintended, split “{skewCol}” into finer categories or add a second tag (region, channel, owner) so decisions are not based on one overloaded label.");
            }

            var trend = ComputeDateActivityTrend(payload, colKinds);
            if (trend == TrendDirection.Up)
            {
                observations.Add("Row activity in the time buckets we detected trends upward toward the end of the period.");
                tips.Add("Double down on what changed in that window (campaigns, pricing, staffing) and document it so you can repeat the lift.");
            }
            else if (trend == TrendDirection.Down)
            {
                observations.Add("Row activity in the detected date buckets softens toward the end of the period.");
                tips.Add("Treat this like an incident: check pipeline, seasonality, stock-outs, or data logging before assuming demand disappeared.");
            }
            else if (trend == TrendDirection.Flat)
            {
                observations.Add("Activity by month looks fairly steady—no dramatic late spike or collapse in the sampled timeline.");
                tips.Add("Use that stability to set a baseline: define “normal” bands so future uploads flag real breaks faster.");
            }

            var dateCol = vm.Columns.FirstOrDefault(c => c.Kind == GenericColumnKind.Date);
            if (dateCol != null && !string.IsNullOrEmpty(dateCol.DateMin) && !string.IsNullOrEmpty(dateCol.DateMax))
                observations.Add(
                    $"Dates span {dateCol.DateMin} through {dateCol.DateMax}. Keep that window in mind when comparing to goals that were set for a different period.");

            if (vm.Columns.All(c => c.Kind != GenericColumnKind.Number))
                tips.Add("Add at least one clean numeric measure (amount, quantity, hours) so the next upload can rank drivers and track improvement.");

            if (vm.Columns.All(c => c.Kind != GenericColumnKind.Date))
                tips.Add("Add a reliable date or timestamp column to separate trend from snapshot and to catch seasonality.");

            if (payload.WasTruncated)
                tips.Add("For production reporting on huge files, move to a database or BI connector so nothing is truncated.");

            tips.Add("Export the profile spreadsheet, share it with the owner of each column, and turn the top three issues into tickets with owners and dates.");

            var teaser = BuildTeaser(worstMissing.Pct, skewShare, trend, topNumeric?.Name, request);
            return new GenericInsightStoryVm
            {
                Teaser = teaser,
                Observations = observations.Distinct().Take(6).ToList(),
                Recommendations = tips.Distinct().Take(7).ToList()
            };
        }

        private enum TrendDirection { Unknown, Up, Down, Flat }

        private static TrendDirection ComputeDateActivityTrend(GenericExcelSessionPayload payload, GenericColumnKind[] colKinds)
        {
            int? dateIdx = null;
            for (var i = 0; i < colKinds.Length; i++)
            {
                if (colKinds[i] != GenericColumnKind.Date) continue;
                dateIdx = i;
                break;
            }
            if (dateIdx is not int dCol) return TrendDirection.Unknown;

            var buckets = new Dictionary<(int y, int m), int>();
            foreach (var row in payload.Rows)
            {
                var cell = dCol < row.Count ? row[dCol] : "";
                if (!TryParseDate(cell, out var dt)) continue;
                var key = (dt.Year, dt.Month);
                buckets[key] = buckets.GetValueOrDefault(key) + 1;
            }
            if (buckets.Count < 4) return TrendDirection.Unknown;

            var ordered = buckets.OrderBy(kv => kv.Key.y).ThenBy(kv => kv.Key.m).Select(kv => kv.Value).ToList();
            var mid = ordered.Count / 2;
            var firstAvg = ordered.Take(mid).DefaultIfEmpty(0).Average();
            var lastAvg = ordered.Skip(mid).DefaultIfEmpty(0).Average();
            if (firstAvg < 1 && lastAvg < 1) return TrendDirection.Unknown;
            if (lastAvg > firstAvg * 1.2) return TrendDirection.Up;
            if (lastAvg < firstAvg * 0.8) return TrendDirection.Down;
            return TrendDirection.Flat;
        }

        private static (string? colName, string label, double share) FindDominantTextCategory(
            GenericExcelSessionPayload payload,
            GenericColumnKind[] colKinds,
            List<GenericColumnSummary> summaries)
        {
            string? bestCol = null;
            var bestLabel = "";
            var bestShare = 0d;
            for (var c = 0; c < colKinds.Length; c++)
            {
                if (colKinds[c] != GenericColumnKind.Text) continue;
                var sum = summaries[c];
                if (sum.DistinctApprox < 2 || sum.DistinctApprox > 80 || sum.NonEmptyCount < 8) continue;
                var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var row in payload.Rows)
                {
                    var v = c < row.Count ? row[c]?.Trim() ?? "" : "";
                    if (v.Length == 0) continue;
                    counts[v] = counts.GetValueOrDefault(v) + 1;
                }
                if (counts.Count == 0) continue;
                var top = counts.OrderByDescending(kv => kv.Value).First();
                var share = top.Value / (double)Math.Max(sum.NonEmptyCount, 1);
                if (share > bestShare)
                {
                    bestShare = share;
                    bestCol = sum.Name;
                    bestLabel = top.Key;
                }
            }
            return (bestCol, bestLabel, bestShare);
        }

        private static string BuildTeaser(double missingPct, double skewShare, TrendDirection trend, string? topNumericName, UserAnalysisRequest? request)
        {
            if (request?.Goal == AnalysisGoalKind.CompareGroups)
                return "Head-to-head view: use the KPIs, chart, and comparison table to see which side is stronger and where to dig next.";
            if (request?.Goal == AnalysisGoalKind.BreakdownByCategory)
                return "Category lens: the tallest bars show where volume or value concentrates—decide whether to lean in or rebalance.";
            if (request?.Goal == AnalysisGoalKind.TrendOverTime)
                return "Timeline lens: the line shows momentum by month—pair it with what the business did in those months.";
            if (missingPct >= 0.35)
                return "The headline is data quality: missing values are bending the charts as much as real performance.";
            if (skewShare >= 0.55)
                return "One label is carrying a lot of the story—great if it is intentional, risky if it hides detail.";
            if (trend == TrendDirection.Down)
                return "Momentum looks cooler toward the end of the period; prioritize finding the leak before debating strategy.";
            if (trend == TrendDirection.Up)
                return "Momentum is warming up—capture what changed while the trail is still fresh.";
            if (!string.IsNullOrEmpty(topNumericName))
                return $"Start with “{topNumericName}”: it is the clearest numeric thread running through this sheet.";
            return "Here is a plain-language take on what this upload is saying, plus concrete ways to act on it.";
        }

        private static string TruncateLabel(string s, int maxLen) =>
            s.Length <= maxLen ? s : s[..(maxLen - 1)] + "…";

        public static string SerializePayload(GenericExcelSessionPayload payload) =>
            JsonSerializer.Serialize(payload);

        public static GenericExcelSessionPayload? DeserializePayload(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try
            {
                return JsonSerializer.Deserialize<GenericExcelSessionPayload>(json);
            }
            catch
            {
                return null;
            }
        }

        private static int DetectHeaderRow(ExcelWorksheet ws, int rowCount, int columnCount)
        {
            var bestRow = 1;
            var bestScore = -1;
            var maxScan = Math.Min(8, rowCount);
            for (int r = 1; r <= maxScan; r++)
            {
                var score = 0;
                for (int c = 1; c <= columnCount; c++)
                {
                    var t = ws.Cells[r, c].Text?.Trim() ?? "";
                    if (t.Length > 0) score++;
                }
                if (score > bestScore)
                {
                    bestScore = score;
                    bestRow = r;
                }
            }
            return bestRow;
        }

        private static bool IsRowEffectivelyEmpty(ExcelWorksheet ws, int r, int colCount)
        {
            for (int c = 1; c <= colCount; c++)
            {
                if (!string.IsNullOrWhiteSpace(ws.Cells[r, c].Text))
                    return false;
            }
            return true;
        }

        private static void DeduplicateHeaders(List<string> headers)
        {
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++)
            {
                var baseName = headers[i];
                if (!seen.ContainsKey(baseName))
                {
                    seen[baseName] = 1;
                    continue;
                }
                seen[baseName]++;
                headers[i] = $"{baseName}_{seen[baseName]}";
            }
        }

        private static GenericColumnKind[] InferColumnKinds(GenericExcelSessionPayload payload)
        {
            var n = payload.Headers.Count;
            var kinds = new GenericColumnKind[n];
            for (int c = 0; c < n; c++)
            {
                int nonEmpty = 0, numOk = 0, dateOk = 0;
                foreach (var row in payload.Rows)
                {
                    var cell = c < row.Count ? row[c] : "";
                    if (string.IsNullOrWhiteSpace(cell)) continue;
                    nonEmpty++;
                    if (TryParseDecimal(cell, out _)) numOk++;
                    if (TryParseDate(cell, out _)) dateOk++;
                }
                if (nonEmpty == 0)
                {
                    kinds[c] = GenericColumnKind.Text;
                    continue;
                }
                var numRatio = (double)numOk / nonEmpty;
                var dateRatio = (double)dateOk / nonEmpty;
                if (numRatio >= 0.72 && numRatio >= dateRatio)
                    kinds[c] = GenericColumnKind.Number;
                else if (dateRatio >= 0.65)
                    kinds[c] = GenericColumnKind.Date;
                else
                    kinds[c] = GenericColumnKind.Text;
            }
            return kinds;
        }

        private static bool TryParseDecimal(string s, out decimal value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out value)) return true;
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;
            return false;
        }

        private static bool TryParseDate(string s, out DateTime value)
        {
            value = default;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out value))
                return true;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out value))
                return true;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var oa))
            {
                try
                {
                    value = DateTime.FromOADate(oa);
                    return true;
                }
                catch { /* ignore */ }
            }
            return false;
        }

        private static string FormatNumber(decimal d) =>
            d == Math.Truncate(d) ? d.ToString("N0", CultureInfo.CurrentCulture) : d.ToString("N2", CultureInfo.CurrentCulture);
    }
}
