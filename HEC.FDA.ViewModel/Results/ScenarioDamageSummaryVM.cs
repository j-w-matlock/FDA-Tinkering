using ClosedXML.Excel;
using HEC.CS.Collections;
using HEC.FDA.Model.metrics;
using HEC.FDA.ViewModel.ImpactAreaScenario;
using HEC.FDA.ViewModel.Saving;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace HEC.FDA.ViewModel.Results
{
    public class ScenarioDamageSummaryVM : BaseViewModel
    {

        public CustomObservableCollection<SelectableChildElement> SelectableElements { get; } = [];
        public CustomObservableCollection<ScenarioDamageRowItem> Rows { get; } = [];
        public CustomObservableCollection<ScenarioPerformanceRowItem> PerformanceRows { get; } = [];
        public CustomObservableCollection<AssuranceOfAEPRowItem> AssuranceOfAEPRows { get; } = [];
        public CustomObservableCollection<ScenarioDamCatRowItem> DamCatRows { get; } = [];

        public ICommand ExportToExcelCommand { get; }

        public ScenarioDamageSummaryVM(List<IASElement> selectedScenarioElems)
        {
            ExportToExcelCommand = new CommandHandler(ExportToExcel, true);

            List<IASElement> allElements = StudyCache.GetChildElementsOfType<IASElement>();

            foreach (IASElement element in allElements)
            {
                SelectableChildElement selectElem = new SelectableChildElement(element);
                SelectableElements.Add(selectElem);
                //the selectable elements are selected by default. We want to toggle all the elements that
                //aren't in the passed in list off.
                if (!selectedScenarioElems.Contains(element))
                {
                    selectElem.IsSelected = false;
                }
                selectElem.SelectionChanged += SelectElem_SelectionChanged;
            }

            LoadTables();
            ListenToChildElementUpdateEvents();
        }

        public ScenarioDamageSummaryVM()
        {
            ExportToExcelCommand = new CommandHandler(ExportToExcel, true);

            List<IASElement> allElements = StudyCache.GetChildElementsOfType<IASElement>();

            foreach (IASElement element in allElements)
            {
                SelectableChildElement selectElem = new SelectableChildElement(element);
                selectElem.SelectionChanged += SelectElem_SelectionChanged;
                SelectableElements.Add(selectElem);
            }

            LoadTables();
            ListenToChildElementUpdateEvents();
        }

        public void ListenToChildElementUpdateEvents()
        {
            StudyCache.IASElementAdded += IASAdded;
            StudyCache.IASElementRemoved += IASRemoved;
            StudyCache.IASElementUpdated += IASUpdated;
        }

        private void IASAdded(object sender, ElementAddedEventArgs e)
        {
            SelectableChildElement newRow = new SelectableChildElement((IASElement)e.Element);
            SelectableElements.Add(newRow);
        }

        private void IASRemoved(object sender, ElementAddedEventArgs e)
        {
            SelectableElements.Remove(SelectableElements.Where(row => row.Element.ID == e.Element.ID).Single());
        }

        private void IASUpdated(object sender, ElementUpdatedEventArgs e)
        {
            IASElement newElement = (IASElement)e.NewElement;
            int idToUpdate = newElement.ID;

            //find the row with this id and update the row's values;
            SelectableChildElement foundRow = SelectableElements.Where(row => row.Element.ID == idToUpdate).SingleOrDefault();
            if (foundRow != null)
            {
                foundRow.Update(newElement);
            }
        }

        private void LoadTables()
        {
            List<IASElement> elems = GetSelectedElements();
            Rows.Clear();
            PerformanceRows.Clear();
            AssuranceOfAEPRows.Clear();
            DamCatRows.Clear();
            foreach (IASElement element in elems)
            {
                Rows.AddRange(ScenarioDamageRowItem.CreateScenarioDamageRowItems(element));
                DamCatRows.AddRange(ScenarioDamCatRowItem.CreateScenarioDamCatRowItems(element));
                List<ImpactAreaScenarioResults> resultsList = element.Results.ResultsList;
                foreach (ImpactAreaScenarioResults impactAreaScenarioResults in resultsList)
                {
                    int iasID = impactAreaScenarioResults.ImpactAreaID;
                    SpecificIAS ias = element.SpecificIASElements.Where(ias => ias.ImpactAreaID == iasID).First();

                    foreach (Threshold threshold in impactAreaScenarioResults.PerformanceByThresholds.ListOfThresholds)
                    {
                        PerformanceRows.Add(new ScenarioPerformanceRowItem(element, ias, threshold));
                        AssuranceOfAEPRows.Add(new AssuranceOfAEPRowItem(element, ias, threshold));
                    }
                }
            }
        }

        private void SelectElem_SelectionChanged(object sender, System.EventArgs e)
        {
            LoadTables();
        }

        private List<IASElement> GetSelectedElements()
        {
            List<IASElement> selectedElements = new List<IASElement>();
            foreach (SelectableChildElement selectElem in SelectableElements)
            {
                if (selectElem.IsSelected)
                {
                    selectedElements.Add(selectElem.Element as IASElement);
                }
            }
            return selectedElements;
        }

        private void ExportToExcel()
        {
            if (!Rows.Any() && !DamCatRows.Any() && !PerformanceRows.Any() && !AssuranceOfAEPRows.Any())
            {
                MessageBox.Show("There are no results available to export.", "Export Scenario Summary Results", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SaveFileDialog dialog = new()
            {
                Title = "Export Scenario Summary Results",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "ScenarioSummaryResults.xlsx",
                AddExtension = true,
                DefaultExt = ".xlsx"
            };

            bool? dialogResult = dialog.ShowDialog();
            if (dialogResult != true || string.IsNullOrWhiteSpace(dialog.FileName))
            {
                return;
            }

            try
            {
                using XLWorkbook workbook = new();
                AddExpectedAnnualDamageTable(workbook);
                AddDamageByCategoryTable(workbook);
                AddPerformanceTable(workbook);
                AddAssuranceOfAepTable(workbook);

                workbook.SaveAs(dialog.FileName);
                MessageBox.Show($"Scenario summary results exported to:{Environment.NewLine}{dialog.FileName}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (IOException ioEx)
            {
                MessageBox.Show($"The export file could not be created. Please close any application using the file and try again.{Environment.NewLine}{ioEx.Message}", "Export Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred while exporting the results.{Environment.NewLine}{ex.Message}", "Export Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddExpectedAnnualDamageTable(XLWorkbook workbook)
        {
            DataTable table = new("Expected Annual Damage");
            table.Columns.Add("Scenario Name");
            table.Columns.Add("Analysis Year");
            table.Columns.Add("Impact Area");
            table.Columns.Add("Mean EAD", typeof(double));
            table.Columns.Add("25th Percentile EAD", typeof(double));
            table.Columns.Add("50th Percentile EAD", typeof(double));
            table.Columns.Add("75th Percentile EAD", typeof(double));

            foreach (ScenarioDamageRowItem row in Rows)
            {
                table.Rows.Add(row.Name, row.AnalysisYear, row.ImpactArea, row.Mean, row.Point75, row.Point5, row.Point25);
            }

            InsertTable(workbook, table, "EAD Summary", "EAD_Summary");
        }

        private void AddDamageByCategoryTable(XLWorkbook workbook)
        {
            DataTable table = new("Expected Annual Damage by Category");
            table.Columns.Add("Scenario Name");
            table.Columns.Add("Analysis Year");
            table.Columns.Add("Impact Area");
            table.Columns.Add("Damage Category");
            table.Columns.Add("Asset Category");
            table.Columns.Add("Mean EAD", typeof(double));

            foreach (ScenarioDamCatRowItem row in DamCatRows)
            {
                table.Rows.Add(row.Name, row.AnalysisYear, row.ImpactAreaName, row.DamCat, row.AssetCat, row.MeanDamage);
            }

            InsertTable(workbook, table, "EAD by Category", "EAD_By_Category");
        }

        private void AddPerformanceTable(XLWorkbook workbook)
        {
            DataTable table = new("Performance Metrics");
            table.Columns.Add("Scenario Name");
            table.Columns.Add("Analysis Year");
            table.Columns.Add("Impact Area");
            table.Columns.Add("Threshold Type");
            table.Columns.Add("Threshold Value", typeof(double));
            table.Columns.Add("Mean AEP", typeof(double));
            table.Columns.Add("Median AEP", typeof(double));
            table.Columns.Add("Long-Term 10", typeof(double));
            table.Columns.Add("Long-Term 30", typeof(double));
            table.Columns.Add("Long-Term 50", typeof(double));
            table.Columns.Add("Assurance 0.1", typeof(double));
            table.Columns.Add("Assurance 0.04", typeof(double));
            table.Columns.Add("Assurance 0.02", typeof(double));
            table.Columns.Add("Assurance 0.01", typeof(double));
            table.Columns.Add("Assurance 0.004", typeof(double));
            table.Columns.Add("Assurance 0.002", typeof(double));

            foreach (ScenarioPerformanceRowItem row in PerformanceRows)
            {
                table.Rows.Add(row.Name, row.AnalysisYear, row.ImpactArea, row.ThresholdType, row.ThresholdValue, row.Mean, row.Median, row.LongTerm10, row.LongTerm30, row.LongTerm50, row.Threshold1, row.Threshold04, row.Threshold02, row.Threshold01, row.Threshold004, row.Threshold002);
            }

            InsertTable(workbook, table, "Performance", "Performance_Metrics");
        }

        private void AddAssuranceOfAepTable(XLWorkbook workbook)
        {
            DataTable table = new("Assurance of AEP");
            table.Columns.Add("Scenario Name");
            table.Columns.Add("Analysis Year");
            table.Columns.Add("Impact Area");
            table.Columns.Add("Threshold Type");
            table.Columns.Add("Threshold Value", typeof(double));
            table.Columns.Add("Mean AEP", typeof(double));
            table.Columns.Add("Median AEP", typeof(double));
            table.Columns.Add("AEP with 90% Assurance", typeof(double));
            table.Columns.Add("AEP 0.1", typeof(double));
            table.Columns.Add("AEP 0.04", typeof(double));
            table.Columns.Add("AEP 0.02", typeof(double));
            table.Columns.Add("AEP 0.01", typeof(double));
            table.Columns.Add("AEP 0.004", typeof(double));
            table.Columns.Add("AEP 0.002", typeof(double));

            foreach (AssuranceOfAEPRowItem row in AssuranceOfAEPRows)
            {
                table.Rows.Add(row.Name, row.AnalysisYear, row.ImpactArea, row.ThresholdType, row.ThresholdValue, row.Mean, row.Median, row.NinetyPercentAssurance, row.AEP1, row.AEP04, row.AEP02, row.AEP01, row.AEP004, row.AEP002);
            }

            InsertTable(workbook, table, "AEP Assurance", "AEP_Assurance");
        }

        private static void InsertTable(XLWorkbook workbook, DataTable table, string worksheetName, string tableName)
        {
            var worksheet = workbook.Worksheets.Add(worksheetName);
            var xlTable = worksheet.Cell(1, 1).InsertTable(table, tableName);
            xlTable.Theme = XLTableTheme.TableStyleMedium2;
            worksheet.Columns().AdjustToContents();
        }
    }
}
