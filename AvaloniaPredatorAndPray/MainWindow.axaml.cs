using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Interactivity;
using LiveChartsCore;
using LiveChartsCore.Defaults;
using LiveChartsCore.Measure;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using OfficeOpenXml;
using SkiaSharp;

namespace AvaloniaPredatorAndPray;

public partial class MainWindow : Window
    {
        private readonly ObservableCollection<ObservablePoint> _preyData = new();
        private readonly ObservableCollection<ObservablePoint> _predatorData = new();

        private readonly ObservableCollection<ObservablePoint>[] _cyclePrey = new ObservableCollection<ObservablePoint>[3];
        private readonly ObservableCollection<ObservablePoint>[] _cyclePredator = new ObservableCollection<ObservablePoint>[3];

        public MainWindow()
        {
            InitializeComponent();
            LoadFileButton.Click += LoadFileButton_Click;
            InitializeCharts();

            for (int i = 0; i < 3; i++)
            {
                _cyclePrey[i] = new ObservableCollection<ObservablePoint>();
                _cyclePredator[i] = new ObservableCollection<ObservablePoint>();
            }
        }

        private void InitializeCharts()
        {
          TimeChart.Series = new ISeries[]
          {
            new LineSeries<ObservablePoint>
            {
                Values = _preyData,
                Name = "Мыши",
                Stroke = new SolidColorPaint(SKColors.Green) { StrokeThickness = 2 },
                Fill = new SolidColorPaint(SKColors.Green), 
                GeometrySize = 5, 
                GeometryStroke = new SolidColorPaint(SKColors.Green), 
                GeometryFill = new SolidColorPaint(SKColors.Green) 
            },
            new LineSeries<ObservablePoint>
            {
                Values = _predatorData,
                Name = "Совы",
                Stroke = new SolidColorPaint(SKColors.Red) { StrokeThickness = 2 },
                Fill = new SolidColorPaint(SKColors.Red), 
                GeometrySize = 5,
                GeometryStroke = new SolidColorPaint(SKColors.Red), 
                GeometryFill = new SolidColorPaint(SKColors.Red) 
            }
          };


            TimeChart.XAxes = new Axis[]
            {
            new Axis
            {
                Labeler = (value) => value.ToString("F1"),
                TextSize = 12,
                Name = "Время"
            }
            };

            TimeChart.YAxes = new Axis[]
            {
            new Axis
            {
                TextSize = 12,
                Name = "Численность"
            }
            };

            TimeChart.LegendPosition = LegendPosition.Right;

            PhaseChart.Series = Array.Empty<ISeries>();
            PhaseChart.XAxes = new Axis[]
            {
            new Axis
            {
                Labeler = (value) => value.ToString("F0"),
                TextSize = 12,
                Name = "Мыши (x)"
            }
            };

            PhaseChart.YAxes = new Axis[]
            {
            new Axis
            {
                TextSize = 12,
                Name = "Совы (y)"
            }
            };

            PhaseChart.LegendPosition = LegendPosition.Right;
        }

        private void UpdatePhaseChart()
        {
            PhaseChart.Series = new ISeries[]
            {
        new LineSeries<ObservablePoint>
        {
            Values = GetPhaseData(_cyclePrey[0], _cyclePredator[0]),
            Name = "Цикл 1",
            Stroke = new SolidColorPaint(SKColors.Blue) { StrokeThickness = 2 },
            GeometrySize = 5,
            GeometryStroke = new SolidColorPaint(SKColors.Blue),
            GeometryFill = new SolidColorPaint(SKColors.Blue)
        },
        new LineSeries<ObservablePoint>
        {
            Values = GetPhaseData(_cyclePrey[1], _cyclePredator[1]),
            Name = "Цикл 2",
            Stroke = new SolidColorPaint(SKColors.Orange) { StrokeThickness = 2 },
            GeometrySize = 5,
            GeometryStroke = new SolidColorPaint(SKColors.Orange),
            GeometryFill = new SolidColorPaint(SKColors.Orange)
        },
        new LineSeries<ObservablePoint>
        {
            Values = GetPhaseData(_cyclePrey[2], _cyclePredator[2]),
            Name = "Цикл 3",
            Stroke = new SolidColorPaint(SKColors.Black) { StrokeThickness = 2 },
            GeometrySize = 5,
            GeometryStroke = new SolidColorPaint(SKColors.Black),
            GeometryFill = new SolidColorPaint(SKColors.Black)
        }
            };
        }


        private async void LoadFileButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Выберите файл Excel"
            };

            dialog.Filters.Add(new FileDialogFilter { Name = "Excel Files", Extensions = { "xlsx" } });

            var result = await dialog.ShowAsync(this);

            if (result != null && result.Any())
            {
                string filePath = result[0];
                try
                {
                    LoadDataFromExcel(filePath);
                    FileInfoText.Text = $"Загружен файл: {Path.GetFileName(filePath)}";
                }
                catch (Exception ex)
                {
                    FileInfoText.Text = $"Ошибка загрузки: {ex.Message}";
                }
            }
        }

        private void LoadDataFromExcel(string filePath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Dasha");
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                double epsilon = GetDoubleValue(worksheet.Cells[4, 2]);
                double alpha = GetDoubleValue(worksheet.Cells[5, 2]);
                double beta = GetDoubleValue(worksheet.Cells[8, 2]);
                double delta = GetDoubleValue(worksheet.Cells[7, 2]);
                double dt = GetDoubleValue(worksheet.Cells[9, 2]);

                EpsilonText.Text = epsilon.ToString("F2");
                AlphaText.Text = alpha.ToString("F3");
                BetaText.Text = beta.ToString("F2");
                DeltaText.Text = delta.ToString("F5");
                DtText.Text = dt.ToString("F2");

                ClearPreviousData();


                ReadCycle(worksheet, 3, 152, _cyclePrey[0], _cyclePredator[0]);
                ReadCycle(worksheet, 153, 302, _cyclePrey[1], _cyclePredator[1]);
                ReadCycle(worksheet, 303, 452, _cyclePrey[2], _cyclePredator[2]);

                foreach (var p in _cyclePrey[0]) _preyData.Add(p);
                foreach (var p in _cyclePredator[0]) _predatorData.Add(p);

                UpdatePhaseChart();
            }
        }

        private void ClearPreviousData()
        {
            _preyData.Clear();
            _predatorData.Clear();
            for (int i = 0; i < 3; i++)
            {
                _cyclePrey[i].Clear();
                _cyclePredator[i].Clear();
            }
        }


        private double GetDoubleValue(ExcelRange cell)
        {
            if (cell.Value == null) return 0;
            return double.TryParse(cell.Value.ToString(), out double result) ? result : 0;
        }

        private void ReadCycle(ExcelWorksheet worksheet, int startRow, int endRow, ObservableCollection<ObservablePoint> preyList, ObservableCollection<ObservablePoint> predatorList)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                if (double.TryParse(worksheet.Cells[row, 3].Value?.ToString(), out double time) &&
                    double.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out double prey) &&
                    double.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out double pred))
                {
                    preyList.Add(new ObservablePoint(time, prey));
                    predatorList.Add(new ObservablePoint(time, pred));
                }
            }
        }

        private ObservableCollection<ObservablePoint> GetPhaseData(ObservableCollection<ObservablePoint> prey, ObservableCollection<ObservablePoint> pred)
        {
            var phase = new ObservableCollection<ObservablePoint>();
            for (int i = 0; i < prey.Count && i < pred.Count; i++)
            {
                phase.Add(new ObservablePoint(prey[i].Y, pred[i].Y));
            }
            return phase;
        }
    }