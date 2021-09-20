using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GenericScenarioEvaluation
{

    public partial class Form1 : Form
    {
        readonly DataSet genScenarios = new DataSet();
        readonly DataTable genericScenarioTable = new DataTable("General Info");
        readonly DataTable ActivityTable = new DataTable("Activity Info");
        readonly DataTable dataValuesTable = new DataTable("Model Info");
        readonly DataTable elementsAndDataTable = new DataTable("Data Element Table");
        readonly DataTable infoTable = new DataTable("GS Details");
        readonly DataTable occExpTable = new DataTable("Occupational Exposure");
        readonly DataTable concentrationTable = new DataTable("Concntrations");
        readonly DataTable calculationTable = new DataTable("Calculations");
        readonly DataTable procDescriptionTable = new DataTable("Process Descriptions");
        readonly DataTable useRateTable = new DataTable("Use Rates");
        readonly DataTable envReleaseTable = new DataTable("Environmental Releases");
        readonly DataTable contolTechTable = new DataTable("Control Technologies");
        readonly DataTable shiftTable = new DataTable("Shifts");
        readonly DataTable operatingDaysTable = new DataTable("Operating Days");
        readonly DataTable workerTable = new DataTable("Workers");
        readonly DataTable siteTable = new DataTable("Number of Sites");
        readonly DataTable ppeTable = new DataTable("PPE");
        readonly DataTable productionRateTable = new DataTable("ProductionRate");
        readonly DataTable parameterTable = new DataTable("Parameters");
        readonly DataTable remainingDataTable = new DataTable("Data Values");
        readonly DataTable sourceTable = new DataTable("Sources");
        readonly DataTable releaseSummaryTable = new DataTable("Release Summary");
        readonly DataTable occExposureSummaryTable = new DataTable("Occupational Exposure Summary");

        readonly DataSet releases = new DataSet();
        readonly DataTable releaseTable = new DataTable("Others");
        readonly DataTable lossFractionReleaseTable = new DataTable("Loss Fraction");
        readonly DataTable throughputReleaseTable = new DataTable("Throughput");
        readonly DataTable ap42ReleaseTable = new DataTable("AP-42");
        readonly DataTable pmnReleaseTable = new DataTable("PMN");
        readonly DataTable opptReleaseTable = new DataTable("OPPT Models");
        readonly DataTable asssumedReleaseTable = new DataTable("Assumed Releases");
        readonly DataTable calculatedReleaseTable = new DataTable("Calculated");
        readonly DataTable agencyReleaseTable = new DataTable("Agency Model");
        readonly DataTable industryReleaseTable = new DataTable("Industry");

        readonly DataSet exposureData = new DataSet();
        readonly DataTable exposureTable = new DataTable("Others");
        readonly DataTable nioshOshaTable = new DataTable("NIOSH-OSHA");
        readonly DataTable pmnExposureTable = new DataTable("PMN");
        readonly DataTable opptExposureTable = new DataTable("OPPT Models");
        readonly DataTable asssumedExpsoureTable = new DataTable("Assumed Releases");
        readonly DataTable calculatedExposureTable = new DataTable("Calculated");
        readonly DataTable agencyExpsoureTable = new DataTable("Agency Model");
        readonly DataTable industryExpsoureTable = new DataTable("Industry");

        readonly List<DataElement> dataElements = new List<DataElement>();
        readonly List<Source> sources = new List<Source>();
        readonly ExposureCollection expsoures = new ExposureCollection();
        readonly List<string> occActivities = new List<string>();

        readonly List<Concentration> concentrations = new List<Concentration>();
        readonly List<Calculation> calculations = new List<Calculation>();
        readonly List<ProcessDescription> processDescriptions = new List<ProcessDescription>();
        readonly List<UseRate> useRates = new List<UseRate>();
        readonly List<EnvironmentalRelease> envRelease = new List<EnvironmentalRelease>();
        readonly List<string> envReleaseActivities = new List<string>();
        readonly List<ControlTechnology> controlTech = new List<ControlTechnology>();
        readonly List<string> controlTechActivities = new List<string>();
        readonly List<Shift> shifts = new List<Shift>();
        readonly List<OperatingDay> opDays = new List<OperatingDay>();
        readonly List<Worker> workers = new List<Worker>();
        readonly List<Site> sites = new List<Site>();
        readonly List<PPE> ppes = new List<PPE>();
        readonly List<ProductionRate> productions = new List<ProductionRate>();
        readonly List<DataValue> values = new List<DataValue>();
        readonly List<string> uniqueDataElements = new List<string>();
        readonly List<string> uniqueDataSubElements = new List<string>();
        readonly List<RemainingValue> remainingValues = new List<RemainingValue>();

        readonly ReleaseCollection cleaningReleases = new ReleaseCollection();
        readonly ReleaseCollection dumpingReleases = new ReleaseCollection();
        readonly ReleaseCollection dryingReleases = new ReleaseCollection();
        readonly ReleaseCollection evaporatingReleases = new ReleaseCollection();
        readonly ReleaseCollection fugitiveReleases = new ReleaseCollection();
        readonly ReleaseCollection disposalReleases = new ReleaseCollection();
        readonly ReleaseCollection residualReleases = new ReleaseCollection();
        readonly ReleaseCollection particulateReleases = new ReleaseCollection();
        readonly ReleaseCollection samplingReleases = new ReleaseCollection();
        readonly ReleaseCollection loadingReleases = new ReleaseCollection();
        readonly ReleaseCollection spentReleases = new ReleaseCollection();
        readonly ReleaseCollection processReleases = new ReleaseCollection();
        readonly ReleaseCollection releaseNotCategorized = new ReleaseCollection();

        readonly DataSet releaseActivities = new DataSet();
        readonly DataTable cleaningReleaseTable = new DataTable("Cleaning");
        readonly DataTable dumpingReleaseTable = new DataTable("Dumping");
        readonly DataTable dryingReleaseTable = new DataTable("Drying");
        readonly DataTable evaporatingReleaseTable = new DataTable("Evaporating");
        readonly DataTable fugitiveReleaseTable = new DataTable("Fugitive");
        readonly DataTable disposalReleaseTable = new DataTable("Disposal");
        readonly DataTable residualReleaseTable = new DataTable("Residual");
        readonly DataTable particulateReleaseTable = new DataTable("Particulate");
        readonly DataTable samplingReleaseTable = new DataTable("Sampling");
        readonly DataTable loadingReleaseTable = new DataTable("Loading");
        readonly DataTable spentReleaseTable = new DataTable("SpentRelease");
        readonly DataTable processReleaseTable = new DataTable("Process");
        readonly DataTable releaseNotCategorizedTable = new DataTable("NotCategorized");
        readonly List<string> rCategorized = new List<string>();

        readonly ExposureCollection cleaningOccExp = new ExposureCollection();
        readonly ExposureCollection dryingOccExp = new ExposureCollection();
        readonly ExposureCollection evaporatingOccExp = new ExposureCollection();
        readonly ExposureCollection dumpingOccExp = new ExposureCollection();
        readonly ExposureCollection fugitiveOccExp = new ExposureCollection();
        readonly ExposureCollection disposalOccExp = new ExposureCollection();
        readonly ExposureCollection residualOccExp = new ExposureCollection();
        readonly ExposureCollection particulateOccExp = new ExposureCollection();
        readonly ExposureCollection samplingOccExp = new ExposureCollection();
        readonly ExposureCollection loadingOccExp = new ExposureCollection();
        readonly ExposureCollection spentOccExp = new ExposureCollection();
        readonly ExposureCollection processOccExp = new ExposureCollection();
        readonly ExposureCollection occExpNotCategorized = new ExposureCollection();

        readonly DataSet exposureActivities = new DataSet();
        readonly DataTable cleaningExposureTable = new DataTable("Cleaning");
        readonly DataTable dumpingExposureTable = new DataTable("Dumping");
        readonly DataTable dryingExposureTable = new DataTable("Drying");
        readonly DataTable evaporatingExposureTable = new DataTable("Evaporating");
        readonly DataTable fugitiveExposureTable = new DataTable("Fugitive");
        readonly DataTable disposalExposureTable = new DataTable("Disposal");
        readonly DataTable residualExposureTable = new DataTable("Residual");
        readonly DataTable particulateExposureTable = new DataTable("Particulate");
        readonly DataTable samplingExposureTable = new DataTable("Sampling");
        readonly DataTable loadingExposureTable = new DataTable("Loading");
        readonly DataTable spentExposureTable = new DataTable("SpentExposure");
        readonly DataTable processExposureTable = new DataTable("Process");
        readonly DataTable expsoureNotCategorizedTable = new DataTable("NotCategorized");

        readonly DataSet epaReview = new DataSet();
        readonly DataTable generalInfo = new DataTable("GeneralInfo");
        readonly DataTable activityInfo = new DataTable("GeneralInfo");
        readonly DataTable equationInfo = new DataTable("GeneralInfo");


        readonly string[] scenarioElements = new string[]{
            "Element #",
            "ESD/GS Name",
            "Data Element",
            "Sub-Element (Type)",
            "Sub-Element (Type 2)",
            "Sub-Element (Exposure Type)",
            "Sub-Element (Activity/Source)",
            "Sub-Element (Media of Release)",
            "Data Element Source Summary",
            "Data Element Source 1",
            "Data Element Source 2",
            "Data Element Source 3",
            "Data Element Source 4",
            "Data Element Source 5",
            "Data Element Source 6",
            "Data Element Source 7",
            "Data Element Source 8",
            "Reviewed",
            "Reference Check"
        };
        private readonly string[] infoColumns = new string[]{
            "Category",
            "Document Type (GS/ESD)",
            "Date Prepared",
            "ESD/GS Name",
            "Full Citation",
            "Developed By",
            "Description",
            "In-paper Industry Descriptor",
            "Industry Code or Description",
            "Industry Code Type"
        };
        private readonly GenericScenario[] scenarios;

        public Form1()
        {
            InitializeComponent();
            SetUpDataTables();
            //ProcessExcel();
            //NullsToString();
            //ExtractSources();
            //AddTablesToSet();
            //CreateElements();
            //scenarios = ProcessScenarios();
            //CategorizeScenarios();
            //CleanUpGSNames();
            //BuildTree();
            //this.propertyGrid1.SelectedObject = this.scenarios;

            ProcessReviews();
        }

        void ProcessReviews()
        {
            foreach (string dir in System.IO.Directory.EnumerateDirectories(@"..\..\Reviewed Scenarios"))
            {
                foreach (string fileName in System.IO.Directory.EnumerateFiles(dir))
                {
                    using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadSheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(fileName, false))
                    {
                        // DataElementsTable
                        DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                        IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>();

                        // Get General Information

                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = sheets.ElementAt(0);
                        string relationshipId = sheet.Id.Value;
                        DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                        DocumentFormat.OpenXml.Spreadsheet.Worksheet workSheet = worksheetPart.Worksheet;
                        DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = workSheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                        IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();
                        string name = string.Empty;
                        string reviewer = string.Empty;
                        string date = string.Empty;
                        DataRow r = generalInfo.NewRow();
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(1))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) reviewer = GetCellValue(spreadSheetDocument, (DocumentFormat.OpenXml.Spreadsheet.Cell)rows.ElementAt(1).ElementAt(2));
                            r["reviewer"] = reviewer;
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(2))
                        {
                            if (cell.CellReference.ToString().StartsWith("D"))
                            {
                                name = GetCellValue(spreadSheetDocument, cell);
                                r["name"] = name;
                            }
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(3))
                        {
                            if (cell.CellReference.ToString().StartsWith("D"))
                            {
                                string temp = GetCellValue(spreadSheetDocument, cell);
                                System.DateTime temp1 = new System.DateTime(1900, 1, 1).AddDays(Double.Parse(temp));
                                date = temp1.Year > 1990 ? temp1.Year.ToString() : temp;
                                r["year"] = date;
                            }
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(4))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["description"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(5))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["flowDiagram"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(6))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["numActvities"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(7))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["numSources"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(8))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["throughput"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(9))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["concCOI"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(10))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["batchSize"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(11))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["batchDuration"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(12))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["batchPerDay"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(13))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["daysOp"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(14))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["NAICS"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(15))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["facSize"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(16))
                        {
                            if (cell.CellReference.ToString().StartsWith("D")) r["MarketShare"] = GetCellValue(spreadSheetDocument, cell);
                        }
                        generalInfo.Rows.Add(r);

                        // Get Activity Information

                        sheet = sheets.ElementAt(1);
                        relationshipId = sheet.Id.Value;
                        worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                        workSheet = worksheetPart.Worksheet;
                        sheetData = workSheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                        rows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();
                        for (int i = 1; i < rows.Count(); i++)
                        {
                            string activity = string.Empty;
                            string chemSteerActivity = string.Empty;
                            string Description = string.Empty;
                            string ExposureType = string.Empty;
                            string exposureValue = string.Empty;
                            string expsoureValueUnits = string.Empty;
                            string modeled = string.Empty;
                            string dataSource = string.Empty;
                            string modelName = string.Empty;
                            string modelReference = string.Empty;
                            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(i))
                            {
                                string reference = cell.CellReference;
                                if (reference.StartsWith("A")) activity = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("B")) chemSteerActivity = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("C")) Description = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("D")) ExposureType = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("E")) exposureValue = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("F")) expsoureValueUnits = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("G")) modeled = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("H")) dataSource = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("I")) modelName = GetCellValue(spreadSheetDocument, cell);
                                if (reference.StartsWith("J")) modelReference = GetCellValue(spreadSheetDocument, cell);
                            }
                            DataRow r1 = activityInfo.NewRow();
                            r1["name"] = name;
                            r1["year"] = date;
                            r1["reviewer"] = reviewer;
                            r1["activity"] = activity;
                            r1["chemSteerActivity"] = chemSteerActivity;
                            r1["Description"] = Description;
                            r1["ExposureType"] = ExposureType;
                            r1["exposureValue"] = exposureValue;
                            r1["expsoureValueUnits"] = expsoureValueUnits;
                            r1["modeled"] = modeled;
                            r1["dataSource"] = dataSource;
                            r1["modelName"] = modelName;
                            r1["modelReference"] = modelReference;
                            activityInfo.Rows.Add(r1);
                        }

                        // Get Activity Information

                        sheet = sheets.ElementAt(2);
                        relationshipId = sheet.Id.Value;
                        worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                        workSheet = worksheetPart.Worksheet;
                        sheetData = workSheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                        rows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();
                        for (int i = 1; i < rows.Count(); i++)
                        {
                            string activity = string.Empty;
                            string equation = string.Empty;
                            string mediaOrRoute = string.Empty;
                            string exposureType = string.Empty;
                            string exposureComponent = string.Empty;
                            string expsoureInputType = string.Empty;
                            string source = string.Empty;
                            string variableDescription = string.Empty;
                            string variableValue = string.Empty;
                            string variableValueUnits = string.Empty;
                            string measuredOrEstimated = string.Empty;
                            string measurementSource = string.Empty;
                            string estimateBasis = string.Empty;
                            string equationUsed = string.Empty;
                            string reference = string.Empty;
                            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(i))
                            {
                                string cellReference = cell.CellReference;
                                if (cellReference.StartsWith("A")) activity = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("B")) equation = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("C")) mediaOrRoute = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("D")) exposureType = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("E")) exposureComponent = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("F")) expsoureInputType = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("G")) source = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("H")) variableDescription = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("I")) variableValue = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("J")) variableValueUnits = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("K")) measuredOrEstimated = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("L")) measurementSource = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("M")) estimateBasis = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("N")) equationUsed = GetCellValue(spreadSheetDocument, cell);
                                if (cellReference.StartsWith("O")) reference = GetCellValue(spreadSheetDocument, cell);
                            }
                            DataRow r1 = equationInfo.NewRow();
                            r1["name"] = name;
                            r1["activity"] = activity;
                            r1["equation"] = equation;
                            r1["mediaOrRoute"] = mediaOrRoute;
                            r1["exposureType"] = exposureType;
                            r1["exposureComponent"] = exposureComponent;
                            r1["source"] = source;
                            r1["variableDescription"] = variableDescription;
                            r1["variableValue"] = variableValue;
                            r1["variableValueUnits"] = variableValueUnits;
                            r1["measuredOrEstimated"] = measuredOrEstimated;
                            r1["measurementSource"] = measurementSource;
                            r1["estimateBasis"] = estimateBasis;
                            r1["equationUsed"] = equationUsed;
                            r1["reference"] = reference;
                            equationInfo.Rows.Add(r1);
                        }
                    }
                }
            }
        }

        void BuildTree()
        {
            foreach (GenericScenario scenario in scenarios)
            {
                TreeNode rootNode = treeView1.Nodes.Add(scenario.ESD_GS_Name, scenario.ESD_GS_Name);
                TreeNode gsNode = rootNode.Nodes.Add("GSInfo", "GSInfo");
                gsNode.Nodes.Add("Category: " + scenario.Category);
                //                TreeNode node = treeView1.Nodes[el.ESD_GS_Name].Nodes.Add(test);
                //if (!string.IsNullOrEmpty(s.SourceSummary)) node.Nodes.Add(de.SourceSummary);
                TreeNode pdNode = rootNode.Nodes.Add("Process Description", "Process Description");
                foreach (ProcessDescription de in scenario.ProcessDescriptions)
                {
                    if (!string.IsNullOrEmpty(de.Activity))
                    {
                        TreeNode deNode = pdNode.Nodes.Add(de.Activity);
                        deNode.Nodes.Add(de.Description);
                    }
                    else pdNode.Nodes.Add(de.Description);
                }
                TreeNode acNode = rootNode.Nodes.Add("Activties", "Activities");
                foreach (Activity ac in scenario.Activities)
                {
                    TreeNode newNode = null;
                    if (acNode.Nodes.ContainsKey(ac.ChemSTEERActivity))
                        newNode = acNode.Nodes.Find(ac.ChemSTEERActivity, false)[0];
                    else newNode = acNode.Nodes.Add(ac.ChemSTEERActivity, ac.ChemSTEERActivity);
                    if (ac.EnvironmentalReleases.Count > 0)
                    {
                        TreeNode erNode = null;
                        if (newNode.Nodes.ContainsKey("Environmental Releases"))
                            erNode = newNode.Nodes.Find("Environmental Releases", false)[0];
                        else erNode = newNode.Nodes.Add("Environmental Releases", "Environmental Releases");
                        foreach (EnvironmentalRelease er in ac.EnvironmentalReleases)
                        {
                            TreeNode node = null;
                            if (erNode.Nodes.ContainsKey(er.MediaOfRelease))
                                node = erNode.Nodes.Find(er.MediaOfRelease, false)[0];
                            else node = erNode.Nodes.Add(er.MediaOfRelease, er.MediaOfRelease);
                            node.Nodes.Add(er.SourceSummary, er.SourceSummary);
                            TreeNode sourceNode = node.Nodes.Add("Sources", "Sources");
                            foreach (Source s in er.sources)
                            {
                                sourceNode.Nodes.Add(s.ReferenceText, s.ReferenceText);
                            }
                        }
                    }
                    if (ac.OccupationalExposures.Count > 0)
                    {
                        TreeNode oeNode = null;
                        if (newNode.Nodes.ContainsKey("Occupational Exposures"))
                            oeNode = newNode.Nodes.Find("Occupational Exposures", false)[0];
                        else oeNode = newNode.Nodes.Add("Occupational Exposures", "Occupational Exposures");
                        foreach (OccupationalExposure oe in ac.OccupationalExposures)
                        {
                            TreeNode node = null;
                            if (oeNode.Nodes.ContainsKey(oe.ExposureType))
                                node = oeNode.Nodes.Find(oe.ExposureType, false)[0];
                            else node = oeNode.Nodes.Add(oe.ExposureType, oe.ExposureType);
                            node.Nodes.Add(oe.SourceSummary, oe.SourceSummary);
                            TreeNode sourceNode = node.Nodes.Add("Sources", "Sources");
                            foreach (Source s in oe.sources)
                            {
                                sourceNode.Nodes.Add(s.ReferenceText, s.ReferenceText);
                            }
                        }
                    }
                }

                if (scenario.Sources.Count > 0)
                {
                    rootNode = rootNode.Nodes.Add("Sources", "Sources");
                    foreach (Source so in scenario.Sources) rootNode.Nodes.Add(so.ReferenceText);
                }
            }
        }

        void CategorizeScenarios()
        {
            int count = 0;
            var elements = from myElement in dataElements.AsEnumerable()
                           where (myElement.ElementName.ToLower().Contains("process description") ||
                           myElement.Type.ToLower().Contains("process description") ||
                           myElement.ElementName.ToLower().Contains("process summary") ||
                           myElement.ElementName.ToLower().Contains("characterization"))
                           select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                el.accessed = true;
                ProcessDescription de = new ProcessDescription()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = Int32.Parse(el.Element),
                    SourceSummary = el.SourceSummary
                };
                if (!string.IsNullOrEmpty(el.Type2))
                {
                    de.Description = el.Type2.TrimStart(); ;
                    de.Activity = el.Type.TrimStart(); ;
                }
                else if (!string.IsNullOrEmpty(el.Type))
                {
                    if (el.Type.Contains(":"))
                    {
                        string[] temp = el.Type.Split(':');
                        de.Activity = temp[0].TrimStart(); ;
                        de.Description = temp[1].TrimStart(); ;
                    }
                    else de.Activity = el.Type;
                }
                else if (!string.IsNullOrEmpty(el.ElementName))
                {
                    if (el.ElementName.Contains(":"))
                    {
                        string[] temp = el.ElementName.Split(':');
                        if (!string.IsNullOrEmpty(temp[1]))
                        {
                            de.Description = temp[1].TrimStart();
                            de.Activity = string.Empty;
                        }
                        else de.Activity = el.Type;
                        if (el.ESD_GS_Name.Contains("PU Foam"))
                        {
                            de.Activity = temp[1];
                            de.Description = String.Empty;
                        }
                    }
                    else if (el.ElementName == "Process Description")
                    {
                        de.Description = el.Type.TrimStart(); ;
                    }
                    else de.Activity = el.ElementName;
                }
                processDescriptions.Add(de);
                de.GenericScenario = GetScenario(el.ESD_GS_Name);
                de.GenericScenario.ProcessDescriptions.Add(de);
                //Activity ac = new Activity() 
                //{

                //};
                //de.GenericScenario.Activities.Add(ac);
                de.sources = GetSources(el);
                procDescriptionTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, de.Activity, de.Description, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.ToLower().Contains("occupational")
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                el.accessed = true;

                OccupationalExposure o = new OccupationalExposure()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = Int32.Parse(el.Element),
                    Type = el.Type,
                    ExposureType = el.ExposureType,
                    ActivitySource = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary
                };

                o.GenericScenario = GetScenario(o.ScenarioName);
                o.GenericScenario.OccupationalExposures.Add(o);
                o.sources = GetSources(el);
                expsoures.Add(o);

                if (el.Activity_Source.ToLower().Contains("unloading"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Unloading Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Unloading Solids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Unloading Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Unloading Liquids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Unloading") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Unloading"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("loading"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Loading Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Loading Solids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Loading Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Loading Liquids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Loading") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Loading"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("equipment") && el.Activity_Source.ToLower().Contains("cleaning"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Equipement Cleaning Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Equipement Cleaning Solids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Equipement Cleaning Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Equipement Cleaning Liquids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Equipement Cleaning") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Equipement Cleaning"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("cleaning"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Cleaning Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Cleaning Solids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Cleaning Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Cleaning Liquids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Cleaning") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Cleaning"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("sampling"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Sampling Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Sampling Solids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Sampling Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Sampling Liquids"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in o.GenericScenario.Activities)
                            if (a.Name == "Sampling") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Sampling"
                            };
                            o.GenericScenario.Activities.Add(ac);
                        }
                        ac.OccupationalExposures.Add(o);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("transfer"))
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Vapor Release from Transfer Operations") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Vapor Release from Transfer Operations"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }
                else if (el.Activity_Source.ToLower().Contains("automotive") && el.Activity_Source.ToLower().Contains("spray") && el.Activity_Source.ToLower().Contains("coating"))
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Automobile Spray Coating") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Automobile Spray Coating"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }
                else if (el.Activity_Source.ToLower().Contains("coating"))
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Generic Coating Applications") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Generic Coating Applications"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }
                else if (el.Activity_Source.ToLower().Contains("plating"))
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Electroplating Bath Additives") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Electroplating Bath Additives"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }
                else if (el.Activity_Source.ToLower().Contains("recirc"))
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Recirculating Water-Cooling Tower Additives") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Recirculating Water-Cooling Tower Additives"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }
                else if (el.Activity_Source.ToLower().Contains("unit") || el.Activity_Source.ToLower().Contains("operation") || el.Activity_Source.ToLower().Contains("process"))
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Unit Operations and Processes") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Unit Operations and Processes"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }
                else
                {
                    Activity ac = null;
                    foreach (Activity a in o.GenericScenario.Activities)
                        if (a.Name == "Miscellaneous Sources/Activities") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Miscellaneous Sources/Activities"
                        };
                        o.GenericScenario.Activities.Add(ac);
                    }
                    ac.OccupationalExposures.Add(o);
                }



                if (el.SourceSummary.ToLower().Contains("niosh") || el.SourceSummary.ToLower().Contains("osha"))
                {
                    nioshOshaTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("pmn"))
                {
                    pmnExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("tec") || el.SourceSummary.ToLower().Contains("industry"))
                {
                    industryExpsoureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("oppt") || el.SourceSummary.ToLower().Contains("ceb"))
                {
                    opptExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("epa") || el.SourceSummary.ToLower().Contains("eu") || el.SourceSummary.ToLower().Contains("oecd") || el.SourceSummary.ToLower().Contains("environment canada"))
                {
                    agencyExpsoureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("calcul"))
                {
                    calculatedExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("assum"))
                {
                    asssumedExpsoureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else
                {
                    exposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }

                occExpTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.ExposureType, o.Dermal ? "1" : "0", o.DermalSolid ? "1" : "0", o.DermalLiquid ? "1" : "0", o.Inhalation ? "1" : "0", o.Particulate ? "1" : "0", o.ChemicalOrVapor ? "1" : "0", el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                if (!occActivities.Contains(o.ActivitySource)) occActivities.Add(o.ActivitySource);
                if (o.ActivitySource.ToLower().Contains("cleaning"))
                {
                    cleaningOccExp.Add(o);
                    cleaningExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("dump"))
                {
                    dumpingOccExp.Add(o);
                    dumpingExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("dry"))
                {
                    dryingOccExp.Add(o);
                    dryingExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("evap"))
                {
                    evaporatingOccExp.Add(o);
                    evaporatingExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("fugit")
                    || (o.ActivitySource.ToLower().Contains("off") && o.ActivitySource.ToLower().Contains("gas"))
                    || o.ActivitySource.ToLower().Contains("emissi")
                    || o.ActivitySource.ToLower().Contains("spill"))
                {
                    fugitiveOccExp.Add(o);
                    fugitiveExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("dispos")
                    || (o.ActivitySource.ToLower().Contains("off") && o.ActivitySource.ToLower().Contains("spec"))
                    || o.ActivitySource.ToLower().Contains("waste"))
                {
                    disposalOccExp.Add(o);
                    disposalExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("resid"))
                {
                    residualOccExp.Add(o);
                    residualExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("partic")
                    || o.ActivitySource.ToLower().Contains("dust"))
                {
                    particulateOccExp.Add(o);
                    particulateExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("sampl"))
                {
                    samplingOccExp.Add(o);
                    samplingExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("loadi")
                    || o.ActivitySource.ToLower().Contains("transf")
                    || o.ActivitySource.ToLower().Contains("handl"))
                {
                    loadingOccExp.Add(o);
                    loadingExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (o.ActivitySource.ToLower().Contains("spent"))
                {
                    spentOccExp.Add(o);
                    spentExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else
                {
                    processOccExp.Add(o);
                    processExposureTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Environmental Release") ||
                       myElement.ElementName.Contains("TRI Releases (lb/yr)") ||
                       myElement.ElementName.Contains("Total Industry Estimated Process Water Discharge Flow"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                el.accessed = true;

                EnvironmentalRelease er = new EnvironmentalRelease()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    Type2 = el.Type2,
                    ActivitySource = el.Activity_Source,
                    MediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary,
                };

                er.GenericScenario = GetScenario(er.ScenarioName);
                er.GenericScenario.EnvironmentalReleases.Add(er);
                er.sources = GetSources(el);
                envRelease.Add(er);


                if (el.Activity_Source.ToLower().Contains("unloading"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Unloading Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Unloading Solids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Unloading Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Unloading Liquids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Unloading") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Unloading"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("loading"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Loading Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Loading Solids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Loading Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Loading Liquids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Loading") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Loading"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("equipment") && el.Activity_Source.ToLower().Contains("cleaning"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Equipement Cleaning Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Equipement Cleaning Solids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Equipement Cleaning Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Equipement Cleaning Liquids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Equipement Cleaning") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Equipement Cleaning"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("cleaning"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Cleaning Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Cleaning Solids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Cleaning Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Cleaning Liquids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Cleaning") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Cleaning"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("sampling"))
                {
                    if (el.Activity_Source.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("solid") || el.ExposureType.ToLower().Contains("partic"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Sampling Solids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Sampling Solids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else if (el.Activity_Source.ToLower().Contains("liquid") || el.ExposureType.ToLower().Contains("liquid"))
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Sampling Liquids") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Sampling Liquids"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                    else
                    {
                        Activity ac = null;
                        foreach (Activity a in er.GenericScenario.Activities)
                            if (a.Name == "Sampling") ac = a;
                        if (ac == null)
                        {
                            ac = new Activity()
                            {
                                Name = el.Activity_Source,
                                ChemSTEERActivity = "Sampling"
                            };
                            er.GenericScenario.Activities.Add(ac);
                        }
                        ac.EnvironmentalReleases.Add(er);
                    }
                }
                else if (el.Activity_Source.ToLower().Contains("transfer"))
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Vapor Release from Transfer Operations") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Vapor Release from Transfer Operations"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }
                else if (el.Activity_Source.ToLower().Contains("automotive") && el.Activity_Source.ToLower().Contains("spray") && el.Activity_Source.ToLower().Contains("coating"))
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Automobile Spray Coating") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Automobile Spray Coating"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }
                else if (el.Activity_Source.ToLower().Contains("coating"))
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Generic Coating Applications") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Generic Coating Applications"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }
                else if (el.Activity_Source.ToLower().Contains("plating"))
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Electroplating Bath Additives") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Electroplating Bath Additives"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }
                else if (el.Activity_Source.ToLower().Contains("recirc"))
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Recirculating Water-Cooling Tower Additives") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Recirculating Water-Cooling Tower Additives"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }
                else if (el.Activity_Source.ToLower().Contains("unit") || el.Activity_Source.ToLower().Contains("operation") || el.Activity_Source.ToLower().Contains("process"))
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Unit Operations and Processes") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Unit Operations and Processes"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }
                else
                {
                    Activity ac = null;
                    foreach (Activity a in er.GenericScenario.Activities)
                        if (a.Name == "Miscellaneous Sources/Activities") ac = a;
                    if (ac == null)
                    {
                        ac = new Activity()
                        {
                            Name = el.Activity_Source,
                            ChemSTEERActivity = "Miscellaneous Sources/Activities"
                        };
                        er.GenericScenario.Activities.Add(ac);
                    }
                    ac.EnvironmentalReleases.Add(er);
                }

                if (el.SourceSummary.ToLower().Contains("pmn"))
                {
                    pmnReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("ilma") || el.SourceSummary.ToLower().Contains("semiconductor") || el.SourceSummary.ToLower().Contains("industry"))
                {
                    industryReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("ap") && el.SourceSummary.ToLower().Contains("42"))
                {
                    ap42ReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("oppt") || (el.SourceSummary.ToLower().Contains("ceb")) && !(el.SourceSummary.ToLower().Contains("loss fraction")))
                {
                    opptReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("epa") || el.SourceSummary.ToLower().Contains("neshap") || el.SourceSummary.ToLower().Contains("eu") || el.SourceSummary.ToLower().Contains("oecd") || el.SourceSummary.ToLower().Contains("environment canada"))
                {
                    agencyReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("loss") || el.SourceSummary.ToLower().Contains("fraction"))
                {
                    lossFractionReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("throughput") || el.SourceSummary.ToLower().Contains("production volume") || el.SourceSummary.ToLower().Contains("pv"))
                {
                    throughputReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("calcul"))
                {
                    calculatedReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (el.SourceSummary.ToLower().Contains("assum"))
                {
                    asssumedReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else
                {
                    releaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                er.GenericScenario = GetScenario(el.ESD_GS_Name);
                er.sources = GetSources(el);
                er.GenericScenario.EnvironmentalReleases.Add(er);
                envRelease.Add(er);
                if (!envReleaseActivities.Contains(er.ActivitySource)) envReleaseActivities.Add(er.ActivitySource);
                if (er.ActivitySource.ToLower().Contains("cleaning"))
                {
                    cleaningReleases.Add(er);
                    cleaningReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("dump"))
                {
                    dumpingReleases.Add(er);
                    dumpingReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("dry"))
                {
                    dryingReleases.Add(er);
                    dryingReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("evap"))
                {
                    evaporatingReleases.Add(er);
                    evaporatingReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("fugit")
                    || (er.ActivitySource.ToLower().Contains("off") && er.ActivitySource.ToLower().Contains("gas"))
                    || er.ActivitySource.ToLower().Contains("emissi")
                    || er.ActivitySource.ToLower().Contains("spill"))
                {
                    fugitiveReleases.Add(er);
                    fugitiveReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("dispos")
                    || (er.ActivitySource.ToLower().Contains("off") && er.ActivitySource.ToLower().Contains("spec"))
                    || er.ActivitySource.ToLower().Contains("waste"))
                {
                    disposalReleases.Add(er);
                    disposalReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("resid"))
                {
                    residualReleases.Add(er);
                    residualReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("partic")
                    || er.ActivitySource.ToLower().Contains("dust"))
                {
                    particulateReleases.Add(er);
                    particulateReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("sampl"))
                {
                    samplingReleases.Add(er);
                    samplingReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("loadi")
                    || er.ActivitySource.ToLower().Contains("transf")
                    || er.ActivitySource.ToLower().Contains("handl"))
                {
                    loadingReleases.Add(er);
                    loadingReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (er.ActivitySource.ToLower().Contains("spent"))
                {
                    spentReleases.Add(er);
                    spentReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else if (!string.IsNullOrEmpty(er.ActivitySource))
                {
                    processReleases.Add(er);
                    processReleaseTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                else
                {
                    releaseNotCategorized.Add(er);
                    releaseNotCategorizedTable.Rows.Add(new string[] { el.ESD_GS_Name, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                }
                envReleaseTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, er.ToAir ? "1" : "0", er.ToLand ? "1" : "0", er.ToWater ? "1" : "0", er.RecycledOrReused ? "1" : "0", er.NotSpecified ? "1" : "0", el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Control Technologies") ||
                       myElement.ElementName.Contains("Treatment Technology"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                el.accessed = true;

                ControlTechnology ct = new ControlTechnology()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ScenarioName = el.ElementName,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                ct.GenericScenario = GetScenario(el.ESD_GS_Name);
                ct.GenericScenario.ControlTechnologies.Add(ct);
                ct.sources = GetSources(el);
                controlTech.Add(ct);
                controlTechActivities.Add(ct.SourceSummary);
                contolTechTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }

            int dv = 0;

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.ToLower().Contains("concentration") ||
                       myElement.ElementName.ToLower().Contains("concentration") ||
                       myElement.Type.ToLower().Contains("concentration") ||
                       myElement.ElementName.ToLower().Contains("fraction") ||
                       myElement.ElementName.ToLower().Contains("percent"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                Concentration o = new Concentration()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = Int32.Parse(el.Element),
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(o.ScenarioName);
                o.GenericScenario.Concentrations.Add(o);
                o.GenericScenario.DataValues.Add(o);
                o.Sources = GetSources(el);
                concentrations.Add(o);
                concentrationTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.ToLower().Contains("calculat") ||
                       myElement.Type.ToLower().Contains("calculat") ||
                       myElement.Type2.ToLower().Contains("calculat") ||
                       myElement.SourceSummary.ToLower().Contains("calculat"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                Calculation o = new Calculation()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = Int32.Parse(el.Element),
                    Type = el.Type,
                    ExposureType = el.ExposureType,
                    Activity = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.GenericScenario.Calculations.Add(o);
                o.GenericScenario.DataValues.Add(o);
                o.Sources = GetSources(el);
                calculations.Add(o);
                calculationTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.ExposureType, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.ToLower().Contains("use rate") ||
                       myElement.Type.ToLower().Contains("use rate") ||
                       myElement.ElementName.ToLower().Contains("daily use") ||
                       myElement.ElementName.ToLower().Contains("annual use"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                UseRate ur = new UseRate()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                useRates.Add(ur);
                ur.GenericScenario = GetScenario(el.ESD_GS_Name);
                ur.GenericScenario.UseRates.Add(ur);
                ur.GenericScenario.DataValues.Add(ur);
                ur.Sources = GetSources(el);
                useRateTable.Rows.Add(new string[] { ur.ElementNumber.ToString(), el.ESD_GS_Name, ur.ElementName, ur.Type, ur.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Shift") ||
                       myElement.Type.Contains("Shift"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                Shift shift = new Shift()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                shift.GenericScenario = GetScenario(el.ESD_GS_Name);
                shift.GenericScenario.Shifts.Add(shift);
                shift.GenericScenario.DataValues.Add(shift);
                shift.Sources = GetSources(el);
                shifts.Add(shift);
                shiftTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }


            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Operating") ||
                       myElement.Type.Contains("Operating"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                OperatingDay day = new OperatingDay()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                day.GenericScenario = GetScenario(el.ESD_GS_Name);
                day.GenericScenario.OperatingDays.Add(day);
                day.GenericScenario.DataValues.Add(day);
                day.Sources = GetSources(el);
                opDays.Add(day);
                operatingDaysTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.ToLower().Contains("worker") ||
                       myElement.ElementName.ToLower().Contains("operator"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                Worker worker = new Worker()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                worker.GenericScenario = GetScenario(el.ESD_GS_Name);
                worker.GenericScenario.Workers.Add(worker);
                worker.GenericScenario.DataValues.Add(worker);
                worker.Sources = GetSources(el);
                workers.Add(worker);
                workerTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
            }


            elements = from myElement in dataElements.AsEnumerable()
                       where ((myElement.ElementName.ToLower().Contains("number") && myElement.ElementName.ToLower().Contains("sites")) ||
                       (myElement.ElementName.ToLower().Contains("domestic") && myElement.ElementName.ToLower().Contains("sites")) ||
                       (myElement.ElementName.ToLower().Contains("number") && myElement.ElementName.ToLower().Contains("plants")) ||
                       (myElement.Type.ToLower().Contains("number") && myElement.Type.ToLower().Contains("sites")) ||
                       (myElement.Type.ToLower().Contains("number") && myElement.Type.ToLower().Contains("plants")))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                Site site = new Site()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ElementName = el.ElementName,
                    ScenarioName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                site.GenericScenario = GetScenario(el.ESD_GS_Name);
                site.Sources = GetSources(el);
                site.GenericScenario.Sites.Add(site);
                site.GenericScenario.DataValues.Add(site);
                sites.Add(site);
                siteTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
            }


            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("PPE") ||
                       myElement.Type.Contains("PPE"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                PPE pp = new PPE()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                pp.GenericScenario = GetScenario(el.ESD_GS_Name);
                pp.GenericScenario.PPEs.Add(pp);
                pp.GenericScenario.DataValues.Add(pp);
                pp.Sources = GetSources(el);
                ppes.Add(pp);
                ppeTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.ToLower().Contains("production rate") ||
                       myElement.ElementName.ToLower().Contains("production volume") ||
                       myElement.ElementName.ToLower().Contains("throughput") ||
                       myElement.ElementName.ToLower().Contains("pv "))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                ProductionRate pp = new ProductionRate()
                {
                    ElementNumber = Int32.Parse(el.Element),
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                pp.GenericScenario = GetScenario(el.ESD_GS_Name);
                pp.GenericScenario.ProductionRates.Add(pp);
                pp.GenericScenario.DataValues.Add(pp);
                pp.Sources = GetSources(el);
                productions.Add(pp);
                productionRateTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where !string.IsNullOrEmpty(myElement.Type)
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                DataValue o = new DataValue()
                {
                    ElementName = el.ElementName,
                    ElementNumber = Int32.Parse(el.Element),
                    ScenarioName = el.ESD_GS_Name,
                    SourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.GenericScenario.Values.Add(o);
                o.GenericScenario.DataValues.Add(o);
                o.Sources = GetSources(el);
                values.Add(o);
                if (!uniqueDataElements.Contains(el.ElementName)) uniqueDataElements.Add(el.ElementName);
                if (!uniqueDataSubElements.Contains(el.Type)) uniqueDataSubElements.Add(el.Type);
                parameterTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }

            elements = from myElement in dataElements.AsEnumerable()
                       where !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                dv++;
                el.accessed = true;

                RemainingValue o = new RemainingValue()
                {
                    ElementName = el.ElementName,
                    ElementNumber = Int32.Parse(el.Element),
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    Type2 = el.Type2,
                    ExposureType = el.ExposureType,
                    Activity_Source = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary
                };
                el.accessed = true;
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.GenericScenario.Parameters.Add(o);
                o.GenericScenario.DataValues.Add(o);
                o.Sources = GetSources(el);
                remainingValues.Add(o);
                if (!uniqueDataElements.Contains(el.ElementName)) uniqueDataElements.Add(el.ElementName);
                if (!uniqueDataSubElements.Contains(el.Type)) uniqueDataSubElements.Add(el.Type);
                remainingDataTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            }

            SetColumnWidths();
            int numrows = 0;
            foreach (DataTable table in genScenarios.Tables)
            {
                numrows += table.Rows.Count;
            }

            envReleaseActivities.Sort();
            occActivities.Sort();



            //this.releaseSummaryTable.Rows.Add(new string[] { "Cleaning", cleaningReleases.ToAir.ToString(), cleaningReleases.ToLand.ToString(), cleaningReleases.ToWater.ToString(), cleaningReleases.NotSpecified.ToString(), cleaningReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Dumping", dumpingReleases.ToAir.ToString(), dumpingReleases.ToLand.ToString(), dumpingReleases.ToWater.ToString(), dumpingReleases.NotSpecified.ToString(), dumpingReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Drying", dryingReleases.ToAir.ToString(), dryingReleases.ToLand.ToString(), dryingReleases.ToWater.ToString(), dryingReleases.NotSpecified.ToString(), dryingReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Evaporating", evaporatingReleases.ToAir.ToString(), evaporatingReleases.ToLand.ToString(), evaporatingReleases.ToWater.ToString(), evaporatingReleases.NotSpecified.ToString(), evaporatingReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Fugitive", fugitiveReleases.ToAir.ToString(), fugitiveReleases.ToLand.ToString(), fugitiveReleases.ToWater.ToString(), fugitiveReleases.NotSpecified.ToString(), fugitiveReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Disposal", disposalReleases.ToAir.ToString(), disposalReleases.ToLand.ToString(), disposalReleases.ToWater.ToString(), disposalReleases.NotSpecified.ToString(), disposalReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Residual", residualReleases.ToAir.ToString(), residualReleases.ToLand.ToString(), residualReleases.ToWater.ToString(), residualReleases.NotSpecified.ToString(), residualReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Particulate", particulateReleases.ToAir.ToString(), particulateReleases.ToLand.ToString(), particulateReleases.ToWater.ToString(), particulateReleases.NotSpecified.ToString(), particulateReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Sampling", samplingReleases.ToAir.ToString(), samplingReleases.ToLand.ToString(), samplingReleases.ToWater.ToString(), samplingReleases.NotSpecified.ToString(), samplingReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Loading", loadingReleases.ToAir.ToString(), loadingReleases.ToLand.ToString(), loadingReleases.ToWater.ToString(), loadingReleases.NotSpecified.ToString(), loadingReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Spent Materials", spentReleases.ToAir.ToString(), spentReleases.ToLand.ToString(), spentReleases.ToWater.ToString(), spentReleases.NotSpecified.ToString(), spentReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Process", processReleases.ToAir.ToString(), processReleases.ToLand.ToString(), processReleases.ToWater.ToString(), processReleases.NotSpecified.ToString(), processReleases.Count.ToString() });
            //this.releaseSummaryTable.Rows.Add(new string[] { "Not Specified", releaseNotCategorized.ToAir.ToString(), releaseNotCategorized.ToLand.ToString(), releaseNotCategorized.ToWater.ToString(), releaseNotCategorized.NotSpecified.ToString(), releaseNotCategorized.Count.ToString() });

            //this.occExposureSummaryTable.Rows.Add(new string[] { "Cleaning", cleaningOccExp.ChemicalOrVapor.ToString(), cleaningOccExp.ParticulateInhalation.ToString(), cleaningOccExp.InhalationNotSpecified.ToString(), cleaningOccExp.TotalInhalation.ToString(),
            //    cleaningOccExp.DermalLiquid.ToString(), cleaningOccExp.DermalSolid.ToString(), cleaningOccExp.DermalNotCategorized.ToString(), cleaningOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Dumping", dumpingOccExp.ChemicalOrVapor.ToString(), dumpingOccExp.ParticulateInhalation.ToString(), dumpingOccExp.InhalationNotSpecified.ToString(), dumpingOccExp.TotalInhalation.ToString(),
            //    dumpingOccExp.DermalLiquid.ToString(), dumpingOccExp.DermalSolid.ToString(), dumpingOccExp.DermalNotCategorized.ToString(), dumpingOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Drying", dryingOccExp.ChemicalOrVapor.ToString(), dryingOccExp.ParticulateInhalation.ToString(), dryingOccExp.InhalationNotSpecified.ToString(), dryingOccExp.TotalInhalation.ToString(),
            //    dryingOccExp.DermalLiquid.ToString(), dryingOccExp.DermalSolid.ToString(), dryingOccExp.DermalNotCategorized.ToString(), dryingOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Evaporating", evaporatingOccExp.ChemicalOrVapor.ToString(), evaporatingOccExp.ParticulateInhalation.ToString(), evaporatingOccExp.InhalationNotSpecified.ToString(), evaporatingOccExp.TotalInhalation.ToString(),
            //    evaporatingOccExp.DermalLiquid.ToString(), evaporatingOccExp.DermalSolid.ToString(), evaporatingOccExp.DermalNotCategorized.ToString(), evaporatingOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Fugitive", fugitiveOccExp.ChemicalOrVapor.ToString(), fugitiveOccExp.ParticulateInhalation.ToString(), fugitiveOccExp.InhalationNotSpecified.ToString(), fugitiveOccExp.TotalInhalation.ToString(),
            //    fugitiveOccExp.DermalLiquid.ToString(), fugitiveOccExp.DermalSolid.ToString(), fugitiveOccExp.DermalNotCategorized.ToString(), fugitiveOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Disposal", disposalOccExp.ChemicalOrVapor.ToString(), disposalOccExp.ParticulateInhalation.ToString(), disposalOccExp.InhalationNotSpecified.ToString(), disposalOccExp.TotalInhalation.ToString(),
            //    disposalOccExp.DermalLiquid.ToString(), disposalOccExp.DermalSolid.ToString(), disposalOccExp.DermalNotCategorized.ToString(), disposalOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Residual", residualOccExp.ChemicalOrVapor.ToString(), residualOccExp.ParticulateInhalation.ToString(), residualOccExp.InhalationNotSpecified.ToString(), residualOccExp.TotalInhalation.ToString(),
            //    residualOccExp.DermalLiquid.ToString(), residualOccExp.DermalSolid.ToString(), residualOccExp.DermalNotCategorized.ToString(), residualOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Particulate", particulateOccExp.ChemicalOrVapor.ToString(), particulateOccExp.ParticulateInhalation.ToString(), particulateOccExp.InhalationNotSpecified.ToString(), particulateOccExp.TotalInhalation.ToString(),
            //    particulateOccExp.DermalLiquid.ToString(), particulateOccExp.DermalSolid.ToString(), particulateOccExp.DermalNotCategorized.ToString(), particulateOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Sampling", samplingOccExp.ChemicalOrVapor.ToString(), samplingOccExp.ParticulateInhalation.ToString(), samplingOccExp.InhalationNotSpecified.ToString(), samplingOccExp.TotalInhalation.ToString(),
            //    samplingOccExp.DermalLiquid.ToString(), samplingOccExp.DermalSolid.ToString(), samplingOccExp.DermalNotCategorized.ToString(), samplingOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Loading", loadingOccExp.ChemicalOrVapor.ToString(), loadingOccExp.ParticulateInhalation.ToString(), loadingOccExp.InhalationNotSpecified.ToString(), loadingOccExp.TotalInhalation.ToString(),
            //    loadingOccExp.DermalLiquid.ToString(), loadingOccExp.DermalSolid.ToString(), loadingOccExp.DermalNotCategorized.ToString(), loadingOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Spent Materials", spentOccExp.ChemicalOrVapor.ToString(), spentOccExp.ParticulateInhalation.ToString(), spentOccExp.InhalationNotSpecified.ToString(), spentOccExp.TotalInhalation.ToString(),
            //    spentOccExp.DermalLiquid.ToString(), spentOccExp.DermalSolid.ToString(), spentOccExp.DermalNotCategorized.ToString(), spentOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Process", processOccExp.ChemicalOrVapor.ToString(), processOccExp.ParticulateInhalation.ToString(), processOccExp.InhalationNotSpecified.ToString(), processOccExp.TotalInhalation.ToString(),
            //    processOccExp.DermalLiquid.ToString(), processOccExp.DermalSolid.ToString(), processOccExp.DermalNotCategorized.ToString(), processOccExp.TotalDermal.ToString()});
            //this.occExposureSummaryTable.Rows.Add(new string[] { "Not Specified", occExpNotCategorized.ChemicalOrVapor.ToString(), occExpNotCategorized.ParticulateInhalation.ToString(), occExpNotCategorized.InhalationNotSpecified.ToString(), occExpNotCategorized.TotalInhalation.ToString(),
            //    occExpNotCategorized.DermalLiquid.ToString(), occExpNotCategorized.DermalSolid.ToString(), occExpNotCategorized.DermalNotCategorized.ToString(), occExpNotCategorized.TotalDermal.ToString()});

            //string output = "Activity\tExposure Type\tModel\tGeneric Scenario\tActitivy Source\tReference" + System.Environment.NewLine;
            //foreach (OccupationalExposure occ in cleaningOccExp)
            //{
            //    output = output + "Cleaning\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in dumpingOccExp)
            //{
            //    output = output + "Dumping\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in dryingOccExp)
            //{
            //    output = output + "Drying\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in evaporatingOccExp)
            //{
            //    output = output + "Evaporating\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in fugitiveOccExp)
            //{
            //    output = output + "Fugitive\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in disposalOccExp)
            //{
            //    output = output + "Disposal\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in residualOccExp)
            //{
            //    output = output + "Residual\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in particulateOccExp)
            //{
            //    output = output + "Particulate\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in samplingOccExp)
            //{
            //    output = output + "Sampling\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in loadingOccExp)
            //{
            //    output = output + "Loading\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in spentOccExp)
            //{
            //    output = output + "Spent Materials\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in processOccExp)
            //{
            //    output = output + "Process\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (OccupationalExposure occ in occExpNotCategorized)
            //{
            //    output = output + "Not Specified\t" + occ.ExposureType + "\t" + occ.sourceSummary + "\t" + occ.ScenarioName + "\t" + occ.ActivitySource + "\t" + (occ.sources.Length > 0 ? occ.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //System.IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\exposures.txt", output);


            //output = "Activity\tMedia of Release\tModel\tGeneric Scenario\tActitivy Source\tReference" + System.Environment.NewLine;
            //foreach (EnvironmentalRelease env in cleaningReleases)
            //{
            //    output = output + "Cleaning\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in dumpingReleases)
            //{
            //    output = output + "Dumping\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in dryingReleases)
            //{
            //    output = output + "Cleaning\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in evaporatingReleases)
            //{
            //    output = output + "Evaporating\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in fugitiveReleases)
            //{
            //    output = output + "Fugitive\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in disposalReleases)
            //{
            //    output = output + "Disposal\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in residualReleases)
            //{
            //    output = output + "Residual\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in particulateReleases)
            //{
            //    output = output + "Particulate\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in samplingReleases)
            //{
            //    output = output + "Sampling\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in loadingReleases)
            //{
            //    output = output + "Loading\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in spentReleases)
            //{
            //    output = output + "Spent Materials\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in processReleases)
            //{
            //    output = output + "Process\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //foreach (EnvironmentalRelease env in releaseNotCategorized)
            //{
            //    output = output + "Not Specified\t" + env.MediaOfRelease + "\t" + env.SourceSummary + "\t" + env.ScenarioName + "\t" + env.ActivitySource + "\t" + (env.sources.Length > 0 ? env.sources[0].ReferenceText : string.Empty) + System.Environment.NewLine;
            //}
            //System.IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\releases.txt", output);

            //string output = "Activity\tChemical Vapor Inhalation\tParticulate Inhalation\tInhalation NotSpecified\tTotal Inhalation\tLiquid Dermal\tSolid Dermal\tDermal Not Categorized\tTotal Dermal\t";
            //output = output + "Release to Air\tReleases to Land\tReleases To Water\tRelease Not Specified\tTotal Releases\n";
            //output = output + "Cleaning \t" + cleaningOccExp.ChemicalOrVapor + "\t" + cleaningOccExp.ParticulateInhalation + "\t" + cleaningOccExp.InhalationNotSpecified + "\t" + cleaningOccExp.TotalInhalation
            //    + "\t" + cleaningOccExp.DermalLiquid + "\t" + cleaningOccExp.DermalSolid + "\t" + cleaningOccExp.DermalNotCategorized + "\t" + cleaningOccExp.TotalDermal + "\t" + cleaningReleases.ToAir
            //    + "\t" + cleaningReleases.ToLand + "\t" + cleaningReleases.ToWater + "\t" + cleaningReleases.NotSpecified + "\t" + cleaningReleases.Count + "\n";
            //output = output + "Dumping \t" + dumpingOccExp.ChemicalOrVapor + "\t" + dumpingOccExp.ParticulateInhalation + "\t" + dumpingOccExp.InhalationNotSpecified + "\t" + +dumpingOccExp.TotalInhalation
            //    + "\t" + dumpingOccExp.DermalLiquid + "\t" + dumpingOccExp.DermalSolid + "\t" + dumpingOccExp.DermalNotCategorized + "\t" + dumpingOccExp.TotalDermal + "\t" + dumpingReleases.ToAir
            //    + "\t" + dumpingReleases.ToLand + "\t" + dumpingReleases.ToWater + "\t" + dumpingReleases.NotSpecified + "\t" + dumpingReleases.Count + "\n";
            //output = output + "Drying \t" + dryingOccExp.ChemicalOrVapor + "\t" + dryingOccExp.ParticulateInhalation + "\t" + dryingOccExp.InhalationNotSpecified + "\t" + +dryingOccExp.TotalInhalation
            //    + "\t" + dryingOccExp.DermalLiquid + "\t" + dryingOccExp.DermalSolid + "\t" + dryingOccExp.DermalNotCategorized + "\t" + dryingOccExp.TotalDermal + "\t" + dryingReleases.ToAir
            //    + "\t" + dryingReleases.ToLand + "\t" + dryingReleases.ToWater + "\t" + dryingReleases.NotSpecified + "\t" + dryingReleases.Count + "\n";
            //output = output + "Evaporating \t" + evaporatingOccExp.ChemicalOrVapor + "\t" + evaporatingOccExp.ParticulateInhalation + "\t" + evaporatingOccExp.InhalationNotSpecified + "\t" + evaporatingOccExp.TotalInhalation
            //    + "\t" + evaporatingOccExp.DermalLiquid + "\t" + evaporatingOccExp.DermalSolid + "\t" + evaporatingOccExp.DermalNotCategorized + "\t" + evaporatingOccExp.TotalDermal + "\t" + evaporatingReleases.ToAir
            //    + "\t" + evaporatingReleases.ToLand + "\t" + evaporatingReleases.ToWater + "\t" + evaporatingReleases.NotSpecified + "\t" + evaporatingReleases.Count + "\n";
            //output = output + "Fugitive \t" + fugitiveOccExp.ChemicalOrVapor + "\t" + fugitiveOccExp.ParticulateInhalation + "\t" + fugitiveOccExp.InhalationNotSpecified + "\t" + fugitiveOccExp.TotalInhalation
            //    + "\t" + fugitiveOccExp.DermalLiquid + "\t" + fugitiveOccExp.DermalSolid + "\t" + fugitiveOccExp.DermalNotCategorized + "\t" + fugitiveOccExp.TotalDermal + "\t" + fugitiveReleases.ToAir
            //    + "\t" + fugitiveReleases.ToLand + "\t" + fugitiveReleases.ToWater + "\t" + fugitiveReleases.NotSpecified + "\t" + fugitiveReleases.Count + "\n";
            //output = output + "Disposal \t" + disposalOccExp.ChemicalOrVapor + "\t" + disposalOccExp.ParticulateInhalation + "\t" + disposalOccExp.InhalationNotSpecified + "\t" + disposalOccExp.TotalInhalation
            //    + "\t" + disposalOccExp.DermalLiquid + "\t" + disposalOccExp.DermalSolid + "\t" + disposalOccExp.DermalNotCategorized + "\t" + disposalOccExp.TotalDermal + "\t" + disposalReleases.ToAir
            //    + "\t" + disposalReleases.ToLand + "\t" + disposalReleases.ToWater + "\t" + disposalReleases.NotSpecified + "\t" + disposalReleases.Count + "\n";
            //output = output + "Residual \t" + residualOccExp.ChemicalOrVapor + "\t" + residualOccExp.ParticulateInhalation + "\t" + residualOccExp.InhalationNotSpecified + "\t" + residualOccExp.TotalInhalation
            //    + "\t" + residualOccExp.DermalLiquid + "\t" + residualOccExp.DermalSolid + "\t" + residualOccExp.DermalNotCategorized + "\t" + residualOccExp.TotalDermal + "\t" + residualReleases.ToAir
            //    + "\t" + residualReleases.ToLand + "\t" + residualReleases.ToWater + "\t" + residualReleases.NotSpecified + "\t" + residualReleases.Count + "\n";
            //output = output + "Particulate \t" + particulateOccExp.ChemicalOrVapor + "\t" + particulateOccExp.ParticulateInhalation + "\t" + particulateOccExp.InhalationNotSpecified + "\t" + particulateOccExp.TotalInhalation
            //    + "\t" + particulateOccExp.DermalLiquid + "\t" + particulateOccExp.DermalSolid + "\t" + particulateOccExp.DermalNotCategorized + "\t" + particulateOccExp.TotalDermal + "\t" + particulateReleases.ToAir
            //    + "\t" + particulateReleases.ToLand + "\t" + particulateReleases.ToWater + "\t" + particulateReleases.NotSpecified + "\t" + particulateReleases.Count + "\n";
            //output = output + "Sampling \t" + samplingOccExp.ChemicalOrVapor + "\t" + samplingOccExp.ParticulateInhalation + "\t" + samplingOccExp.InhalationNotSpecified + "\t" + samplingOccExp.TotalInhalation
            //    + "\t" + samplingOccExp.DermalLiquid + "\t" + samplingOccExp.DermalSolid + "\t" + samplingOccExp.DermalNotCategorized + "\t" + samplingOccExp.TotalDermal + "\t" + samplingReleases.ToAir
            //    + "\t" + samplingReleases.ToLand + "\t" + samplingReleases.ToWater + "\t" + samplingReleases.NotSpecified + "\t" + samplingReleases.Count + "\n";
            //output = output + "Loading \t" + loadingOccExp.ChemicalOrVapor + "\t" + loadingOccExp.ParticulateInhalation + "\t" + loadingOccExp.InhalationNotSpecified + "\t" + loadingOccExp.TotalInhalation
            //    + "\t" + loadingOccExp.DermalLiquid + "\t" + loadingOccExp.DermalSolid + "\t" + loadingOccExp.DermalNotCategorized + "\t" + loadingOccExp.TotalDermal + "\t" + loadingReleases.ToAir
            //    + "\t" + loadingReleases.ToLand + "\t" + loadingReleases.ToWater + "\t" + loadingReleases.NotSpecified + "\t" + loadingReleases.Count + "\n";
            //output = output + "Spent Materials \t" + spentOccExp.ChemicalOrVapor + "\t" + spentOccExp.ParticulateInhalation + "\t" + spentOccExp.InhalationNotSpecified + "\t" + spentOccExp.TotalInhalation
            //    + "\t" + spentOccExp.DermalLiquid + "\t" + spentOccExp.DermalSolid + "\t" + spentOccExp.DermalNotCategorized + "\t" + spentOccExp.TotalDermal + "\t" + spentReleases.ToAir
            //    + "\t" + spentReleases.ToLand + "\t" + spentReleases.ToWater + "\t" + spentReleases.NotSpecified + "\t" + spentReleases.Count + "\n";
            //output = output + "Process \t" + processOccExp.ChemicalOrVapor + "\t" + processOccExp.ParticulateInhalation + "\t" + processOccExp.InhalationNotSpecified + "\t" + processOccExp.TotalInhalation
            //    + "\t" + processOccExp.DermalLiquid + "\t" + processOccExp.DermalSolid + "\t" + processOccExp.DermalNotCategorized + "\t" + processOccExp.TotalDermal + "\t" + processReleases.ToAir
            //    + "\t" + processReleases.ToLand + "\t" + processReleases.ToWater + "\t" + processReleases.NotSpecified + "\t" + cleaningReleases.Count + "\n";
            //output = output + "Not Categorized \t" + occExpNotCategorized.ChemicalOrVapor + "\t" + occExpNotCategorized.ParticulateInhalation + "\t" + occExpNotCategorized.InhalationNotSpecified + "\t" + occExpNotCategorized.TotalInhalation
            //    + "\t" + occExpNotCategorized.DermalLiquid + "\t" + occExpNotCategorized.DermalSolid + "\t" + occExpNotCategorized.DermalNotCategorized + "\t" + occExpNotCategorized.TotalDermal + "\t" + releaseNotCategorized.ToAir + "\t" + releaseNotCategorized.ToLand
            //    + "\t" + releaseNotCategorized.ToWater + "\t" + releaseNotCategorized.NotSpecified + "\t" + releaseNotCategorized.Count + "\n";
            //System.IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\table.txt", output);
            this.ProcessGenericScenarios(scenarios);
            try
            {
                ExportDataSet(genScenarios, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\GenericScenarioOutputs.xlsx");
                ExportDataSet(releases, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ReleasesOutputs.xlsx");
                ExportDataSet(exposureData, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ExposureOutputs.xlsx");
                ExportDataSet(releaseActivities, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ReleasesActivities.xlsx");
                ExportDataSet(exposureActivities, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ExposureActivities.xlsx");
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        void SetUpDataTables()
        {

            // General Info from EPA Review
            this.generalInfo.Columns.Add("reviewer");
            this.generalInfo.Columns.Add("name");
            this.generalInfo.Columns.Add("year");
            this.generalInfo.Columns.Add("description");
            this.generalInfo.Columns.Add("flowDiagram");
            this.generalInfo.Columns.Add("numActvities");
            this.generalInfo.Columns.Add("numSources");
            this.generalInfo.Columns.Add("throughput");
            this.generalInfo.Columns.Add("concCOI");
            this.generalInfo.Columns.Add("batchSize");
            this.generalInfo.Columns.Add("batchDuration");
            this.generalInfo.Columns.Add("batchPerDay");
            this.generalInfo.Columns.Add("daysOp");
            this.generalInfo.Columns.Add("NAICS");
            this.generalInfo.Columns.Add("facSize");
            this.generalInfo.Columns.Add("MarketShare");

            // Activity Info from EPA review
            this.activityInfo.Columns.Add("name");
            this.activityInfo.Columns.Add("reviewer");
            this.activityInfo.Columns.Add("year");
            this.activityInfo.Columns.Add("activity");
            this.activityInfo.Columns.Add("chemSteerActivity");
            this.activityInfo.Columns.Add("Description");
            this.activityInfo.Columns.Add("ExposureType");
            this.activityInfo.Columns.Add("exposureValue");
            this.activityInfo.Columns.Add("expsoureValueUnits");
            this.activityInfo.Columns.Add("modeled");
            this.activityInfo.Columns.Add("dataSource");
            this.activityInfo.Columns.Add("modelName");
            this.activityInfo.Columns.Add("modelReference");

            // EquationInfo from EPA Review

            this.equationInfo.Columns.Add("name");
            this.equationInfo.Columns.Add("activity");
            this.equationInfo.Columns.Add("equation");
            this.equationInfo.Columns.Add("mediaOrRoute");
            this.equationInfo.Columns.Add("exposureType");
            this.equationInfo.Columns.Add("exposureComponent");
            this.equationInfo.Columns.Add("source");
            this.equationInfo.Columns.Add("variableDescription");
            this.equationInfo.Columns.Add("variableValue");
            this.equationInfo.Columns.Add("variableValueUnits");
            this.equationInfo.Columns.Add("measuredOrEstimated");
            this.equationInfo.Columns.Add("measurementSource");
            this.equationInfo.Columns.Add("estimateBasis");
            this.equationInfo.Columns.Add("equationUsed");
            this.equationInfo.Columns.Add("reference");


            // dataValueDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type,el.SourceSummary });
            this.parameterTable.Columns.Add("Element Number");
            this.parameterTable.Columns.Add("Scenario Name");
            this.parameterTable.Columns.Add("Element Name");
            this.parameterTable.Columns.Add("Element Type");
            this.parameterTable.Columns.Add("Element Type2");
            this.parameterTable.Columns.Add("Source Summary");

            // remainingValuesDataGridView1.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            this.remainingDataTable.Columns.Add("Element Number");
            this.remainingDataTable.Columns.Add("Scenario Name");
            this.remainingDataTable.Columns.Add("Element Name");
            this.remainingDataTable.Columns.Add("Activity Source");
            this.remainingDataTable.Columns.Add("Media Of Release");
            this.remainingDataTable.Columns.Add("Source Summary");


            // occupationalExposureDataGridView.Rows.Add(new string[] { o.ElementNumber, o.ScenarioName, o.ElementNumber, o.Type, o.Activity_Source, o.sourceSummary, o.mediaOfRelease, o.sourceSummary });
            this.occExpTable.Columns.Add("Element Number");
            this.occExpTable.Columns.Add("Scenario Name");
            this.occExpTable.Columns.Add("Element Name");
            this.occExpTable.Columns.Add("Element Type");
            this.occExpTable.Columns.Add("Exposure Type");
            this.occExpTable.Columns.Add("Dermal");
            this.occExpTable.Columns.Add("Dermal Solid");
            this.occExpTable.Columns.Add("Dermal Liquid");
            this.occExpTable.Columns.Add("Inhalation");
            this.occExpTable.Columns.Add("Inhalation Particulate");
            this.occExpTable.Columns.Add("Inhalation Chemical or Vapor");
            this.occExpTable.Columns.Add("Activity Source");
            this.occExpTable.Columns.Add("Media Of Release");
            this.occExpTable.Columns.Add("Source Summary");

            this.occExposureSummaryTable.Columns.Add(new DataColumn("Activity"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Chemical Vapor Inhalation"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Particulate Inhalation"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Inhalation Not Specified"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Total Inhalation"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Liquid Dermal"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Solid Dermal"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Dermal Not Categorized"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Total Dermal"));

            // occupationalExposureDataGridView.Rows.Add(new string[] { o.ElementNumber, o.ScenarioName, o.ElementNumber, o.Type, o.Activity_Source, o.sourceSummary, o.mediaOfRelease, o.sourceSummary });
            this.concentrationTable.Columns.Add("Element Number");
            this.concentrationTable.Columns.Add("Scenario Name");
            this.concentrationTable.Columns.Add("Element Name");
            this.concentrationTable.Columns.Add("Element Type");
            this.concentrationTable.Columns.Add("Source Summary");

            // occupationalExposureDataGridView.Rows.Add(new string[] { o.ElementNumber, o.ScenarioName, o.ElementNumber, o.Type, o.Activity_Source, o.sourceSummary, o.mediaOfRelease, o.sourceSummary });
            this.calculationTable.Columns.Add("Element Number");
            this.calculationTable.Columns.Add("Scenario Name");
            this.calculationTable.Columns.Add("Element Name");
            this.calculationTable.Columns.Add("Element Type");
            this.calculationTable.Columns.Add("Exposure Type");
            this.calculationTable.Columns.Add("Activity Source");
            this.calculationTable.Columns.Add("Media Of Release");
            this.calculationTable.Columns.Add("Source Summary");

            // processDescriptionDataGridView.Rows.Add(new string[] { de.ElementNumber, de.ElementName, de.Type, de.Type2, de.SourceSummary
            this.procDescriptionTable.Columns.Add("Element Number");
            this.procDescriptionTable.Columns.Add("Scenario Name");
            this.procDescriptionTable.Columns.Add("Activity");
            this.procDescriptionTable.Columns.Add("Description");
            this.procDescriptionTable.Columns.Add("Source Summary");

            // useRateDataGridView.Rows.Add(new string[] { ur.ElementNumber, ur.ElementName, ur.Type, ur.SourceSummary });
            this.useRateTable.Columns.Add("Element Number");
            this.useRateTable.Columns.Add("Scenario Name");
            this.useRateTable.Columns.Add("Element Name");
            this.useRateTable.Columns.Add("Element Type");
            this.useRateTable.Columns.Add("Source Summary");

            // environmentalReleaseDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            this.envReleaseTable.Columns.Add("Element Number");
            this.envReleaseTable.Columns.Add("Scenario Name");
            this.envReleaseTable.Columns.Add("Element Name");
            this.envReleaseTable.Columns.Add("Element Type");
            this.envReleaseTable.Columns.Add("Element Type 2");
            this.envReleaseTable.Columns.Add("Activity_Source");
            this.envReleaseTable.Columns.Add("Media Of Release");
            this.envReleaseTable.Columns.Add("To Air");
            this.envReleaseTable.Columns.Add("To Land");
            this.envReleaseTable.Columns.Add("To Water");
            this.envReleaseTable.Columns.Add("Recycled or Reused");
            this.envReleaseTable.Columns.Add("Not Specified");
            this.envReleaseTable.Columns.Add("Source Summary");

            this.releaseSummaryTable.Columns.Add(new DataColumn("Activity"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release to Air"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release to Land"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release to Water"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release Not Specified"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Total Releases"));

            // controlTechnologyDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            this.contolTechTable.Columns.Add("Element Number");
            this.contolTechTable.Columns.Add("Scenario Name");
            this.contolTechTable.Columns.Add("Element Name");
            this.contolTechTable.Columns.Add("Element Type");
            this.contolTechTable.Columns.Add("Element Type 2");
            this.contolTechTable.Columns.Add("Source Summary");

            // shiftDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            this.shiftTable.Columns.Add("Element Number");
            this.shiftTable.Columns.Add("Scenario Name");
            this.shiftTable.Columns.Add("Element Name");
            this.shiftTable.Columns.Add("Element Type");
            this.shiftTable.Columns.Add("Element Type 2");
            this.shiftTable.Columns.Add("Source Summary");

            // operatingDaysDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            this.operatingDaysTable.Columns.Add("Element Number");
            this.operatingDaysTable.Columns.Add("Scenario Name");
            this.operatingDaysTable.Columns.Add("Element Name");
            this.operatingDaysTable.Columns.Add("Element Type");
            this.operatingDaysTable.Columns.Add("Element Type 2");
            this.operatingDaysTable.Columns.Add("Source Summary");

            // workersDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.workerTable.Columns.Add("Element Number");
            this.workerTable.Columns.Add("Scenario Name");
            this.workerTable.Columns.Add("Element Name");
            this.workerTable.Columns.Add("Element Type");
            this.workerTable.Columns.Add("Source Summary");

            // sitesDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.siteTable.Columns.Add("Element Number");
            this.siteTable.Columns.Add("Scenario Name");
            this.siteTable.Columns.Add("Element Name");
            this.siteTable.Columns.Add("Element Type");
            this.siteTable.Columns.Add("Source Summary");

            // ppeDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.ppeTable.Columns.Add("Element Number");
            this.ppeTable.Columns.Add("Scenario Name");
            this.ppeTable.Columns.Add("Element Name");
            this.ppeTable.Columns.Add("Element Type");
            this.ppeTable.Columns.Add("Source Summary");

            // productionRateDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.productionRateTable.Columns.Add("Element Number");
            this.productionRateTable.Columns.Add("Scenario Name");
            this.productionRateTable.Columns.Add("Element Name");
            this.productionRateTable.Columns.Add("Element Type");
            this.productionRateTable.Columns.Add("Source Summary");

            this.releaseTable.Columns.Add("Generic Scenario");
            this.releaseTable.Columns.Add("Activity");
            this.releaseTable.Columns.Add("Media");
            this.releaseTable.Columns.Add("Summary");

            this.lossFractionReleaseTable.Columns.Add("Generic Scenario");
            this.lossFractionReleaseTable.Columns.Add("Activity");
            this.lossFractionReleaseTable.Columns.Add("Media");
            this.lossFractionReleaseTable.Columns.Add("Summary");

            this.throughputReleaseTable.Columns.Add("Generic Scenario");
            this.throughputReleaseTable.Columns.Add("Activity");
            this.throughputReleaseTable.Columns.Add("Media");
            this.throughputReleaseTable.Columns.Add("Summary");

            this.ap42ReleaseTable.Columns.Add("Generic Scenario");
            this.ap42ReleaseTable.Columns.Add("Activity");
            this.ap42ReleaseTable.Columns.Add("Media");
            this.ap42ReleaseTable.Columns.Add("Summary");

            this.pmnReleaseTable.Columns.Add("Generic Scenario");
            this.pmnReleaseTable.Columns.Add("Activity");
            this.pmnReleaseTable.Columns.Add("Media");
            this.pmnReleaseTable.Columns.Add("Summary");

            this.opptReleaseTable.Columns.Add("Generic Scenario");
            this.opptReleaseTable.Columns.Add("Activity");
            this.opptReleaseTable.Columns.Add("Media");
            this.opptReleaseTable.Columns.Add("Summary");

            this.asssumedReleaseTable.Columns.Add("Generic Scenario");
            this.asssumedReleaseTable.Columns.Add("Activity");
            this.asssumedReleaseTable.Columns.Add("Media");
            this.asssumedReleaseTable.Columns.Add("Summary");

            this.calculatedReleaseTable.Columns.Add("Generic Scenario");
            this.calculatedReleaseTable.Columns.Add("Activity");
            this.calculatedReleaseTable.Columns.Add("Media");
            this.calculatedReleaseTable.Columns.Add("Summary");

            this.agencyReleaseTable.Columns.Add("Generic Scenario");
            this.agencyReleaseTable.Columns.Add("Activity");
            this.agencyReleaseTable.Columns.Add("Media");
            this.agencyReleaseTable.Columns.Add("Summary");

            this.industryReleaseTable.Columns.Add("Generic Scenario");
            this.industryReleaseTable.Columns.Add("Activity");
            this.industryReleaseTable.Columns.Add("Media");
            this.industryReleaseTable.Columns.Add("Summary");

            this.exposureTable.Columns.Add("Generic Scenario");
            this.exposureTable.Columns.Add("Activity");
            this.exposureTable.Columns.Add("Type");
            this.exposureTable.Columns.Add("Summary");

            this.nioshOshaTable.Columns.Add("Generic Scenario");
            this.nioshOshaTable.Columns.Add("Activity");
            this.nioshOshaTable.Columns.Add("Type");
            this.nioshOshaTable.Columns.Add("Summary");

            this.pmnExposureTable.Columns.Add("Generic Scenario");
            this.pmnExposureTable.Columns.Add("Activity");
            this.pmnExposureTable.Columns.Add("Type");
            this.pmnExposureTable.Columns.Add("Summary");

            this.opptExposureTable.Columns.Add("Generic Scenario");
            this.opptExposureTable.Columns.Add("Activity");
            this.opptExposureTable.Columns.Add("Type");
            this.opptExposureTable.Columns.Add("Summary");

            this.asssumedExpsoureTable.Columns.Add("Generic Scenario");
            this.asssumedExpsoureTable.Columns.Add("Activity");
            this.asssumedExpsoureTable.Columns.Add("Type");
            this.asssumedExpsoureTable.Columns.Add("Summary");

            this.calculatedExposureTable.Columns.Add("Generic Scenario");
            this.calculatedExposureTable.Columns.Add("Activity");
            this.calculatedExposureTable.Columns.Add("Type");
            this.calculatedExposureTable.Columns.Add("Summary");

            this.agencyExpsoureTable.Columns.Add("Generic Scenario");
            this.agencyExpsoureTable.Columns.Add("Activity");
            this.agencyExpsoureTable.Columns.Add("Type");
            this.agencyExpsoureTable.Columns.Add("Summary");

            this.industryExpsoureTable.Columns.Add("Generic Scenario");
            this.industryExpsoureTable.Columns.Add("Activity");
            this.industryExpsoureTable.Columns.Add("Type");
            this.industryExpsoureTable.Columns.Add("Summary");

            this.cleaningReleaseTable.Columns.Add("Generic Scenario");
            this.cleaningReleaseTable.Columns.Add("Activity");
            this.cleaningReleaseTable.Columns.Add("Type");
            this.cleaningReleaseTable.Columns.Add("Summary");

            this.dumpingReleaseTable.Columns.Add("Generic Scenario");
            this.dumpingReleaseTable.Columns.Add("Activity");
            this.dumpingReleaseTable.Columns.Add("Type");
            this.dumpingReleaseTable.Columns.Add("Summary");

            this.dryingReleaseTable.Columns.Add("Generic Scenario");
            this.dryingReleaseTable.Columns.Add("Activity");
            this.dryingReleaseTable.Columns.Add("Type");
            this.dryingReleaseTable.Columns.Add("Summary");

            this.evaporatingReleaseTable.Columns.Add("Generic Scenario");
            this.evaporatingReleaseTable.Columns.Add("Activity");
            this.evaporatingReleaseTable.Columns.Add("Type");
            this.evaporatingReleaseTable.Columns.Add("Summary");

            this.fugitiveReleaseTable.Columns.Add("Generic Scenario");
            this.fugitiveReleaseTable.Columns.Add("Activity");
            this.fugitiveReleaseTable.Columns.Add("Type");
            this.fugitiveReleaseTable.Columns.Add("Summary");

            this.disposalReleaseTable.Columns.Add("Generic Scenario");
            this.disposalReleaseTable.Columns.Add("Activity");
            this.disposalReleaseTable.Columns.Add("Type");
            this.disposalReleaseTable.Columns.Add("Summary");

            this.residualReleaseTable.Columns.Add("Generic Scenario");
            this.residualReleaseTable.Columns.Add("Activity");
            this.residualReleaseTable.Columns.Add("Type");
            this.residualReleaseTable.Columns.Add("Summary");

            this.particulateReleaseTable.Columns.Add("Generic Scenario");
            this.particulateReleaseTable.Columns.Add("Activity");
            this.particulateReleaseTable.Columns.Add("Type");
            this.particulateReleaseTable.Columns.Add("Summary");

            this.samplingReleaseTable.Columns.Add("Generic Scenario");
            this.samplingReleaseTable.Columns.Add("Activity");
            this.samplingReleaseTable.Columns.Add("Type");
            this.samplingReleaseTable.Columns.Add("Summary");

            this.loadingReleaseTable.Columns.Add("Generic Scenario");
            this.loadingReleaseTable.Columns.Add("Activity");
            this.loadingReleaseTable.Columns.Add("Type");
            this.loadingReleaseTable.Columns.Add("Summary");

            this.spentReleaseTable.Columns.Add("Generic Scenario");
            this.spentReleaseTable.Columns.Add("Activity");
            this.spentReleaseTable.Columns.Add("Type");
            this.spentReleaseTable.Columns.Add("Summary");

            this.processReleaseTable.Columns.Add("Generic Scenario");
            this.processReleaseTable.Columns.Add("Activity");
            this.processReleaseTable.Columns.Add("Type");
            this.processReleaseTable.Columns.Add("Summary");

            this.releaseNotCategorizedTable.Columns.Add("Generic Scenario");
            this.releaseNotCategorizedTable.Columns.Add("Activity");
            this.releaseNotCategorizedTable.Columns.Add("Type");
            this.releaseNotCategorizedTable.Columns.Add("Summary");

            this.cleaningExposureTable.Columns.Add("Generic Scenario");
            this.cleaningExposureTable.Columns.Add("Activity");
            this.cleaningExposureTable.Columns.Add("Type");
            this.cleaningExposureTable.Columns.Add("Summary");

            this.dumpingExposureTable.Columns.Add("Generic Scenario");
            this.dumpingExposureTable.Columns.Add("Activity");
            this.dumpingExposureTable.Columns.Add("Type");
            this.dumpingExposureTable.Columns.Add("Summary");

            this.dryingExposureTable.Columns.Add("Generic Scenario");
            this.dryingExposureTable.Columns.Add("Activity");
            this.dryingExposureTable.Columns.Add("Type");
            this.dryingExposureTable.Columns.Add("Summary");

            this.evaporatingExposureTable.Columns.Add("Generic Scenario");
            this.evaporatingExposureTable.Columns.Add("Activity");
            this.evaporatingExposureTable.Columns.Add("Type");
            this.evaporatingExposureTable.Columns.Add("Summary");

            this.fugitiveExposureTable.Columns.Add("Generic Scenario");
            this.fugitiveExposureTable.Columns.Add("Activity");
            this.fugitiveExposureTable.Columns.Add("Type");
            this.fugitiveExposureTable.Columns.Add("Summary");

            this.disposalExposureTable.Columns.Add("Generic Scenario");
            this.disposalExposureTable.Columns.Add("Activity");
            this.disposalExposureTable.Columns.Add("Type");
            this.disposalExposureTable.Columns.Add("Summary");

            this.residualExposureTable.Columns.Add("Generic Scenario");
            this.residualExposureTable.Columns.Add("Activity");
            this.residualExposureTable.Columns.Add("Type");
            this.residualExposureTable.Columns.Add("Summary");

            this.particulateExposureTable.Columns.Add("Generic Scenario");
            this.particulateExposureTable.Columns.Add("Activity");
            this.particulateExposureTable.Columns.Add("Type");
            this.particulateExposureTable.Columns.Add("Summary");

            this.samplingExposureTable.Columns.Add("Generic Scenario");
            this.samplingExposureTable.Columns.Add("Activity");
            this.samplingExposureTable.Columns.Add("Type");
            this.samplingExposureTable.Columns.Add("Summary");

            this.loadingExposureTable.Columns.Add("Generic Scenario");
            this.loadingExposureTable.Columns.Add("Activity");
            this.loadingExposureTable.Columns.Add("Type");
            this.loadingExposureTable.Columns.Add("Summary");

            this.spentExposureTable.Columns.Add("Generic Scenario");
            this.spentExposureTable.Columns.Add("Activity");
            this.spentExposureTable.Columns.Add("Type");
            this.spentExposureTable.Columns.Add("Summary");

            this.processExposureTable.Columns.Add("Generic Scenario");
            this.processExposureTable.Columns.Add("Activity");
            this.processExposureTable.Columns.Add("Type");
            this.processExposureTable.Columns.Add("Summary");

            this.expsoureNotCategorizedTable.Columns.Add("Generic Scenario");
            this.expsoureNotCategorizedTable.Columns.Add("Activity");
            this.expsoureNotCategorizedTable.Columns.Add("Type");
            this.expsoureNotCategorizedTable.Columns.Add("Summary");

            generalInfoDataGridView.DataSource = generalInfo;
            activityInfoDataGridView.DataSource = activityInfo;
            equationInfoDataGridView.DataSource = equationInfo;
            processDescriptionDataGridView.DataSource = procDescriptionTable;
            occupationalExposureDataGridView.DataSource = occExpTable;
            environmentalReleaseDataGridView.DataSource = envReleaseTable;
            controlTechnologyDataGridView.DataSource = contolTechTable;
            ConcentrationDataGridView.DataSource = concentrationTable;
            CalculationdataGridView.DataSource = calculationTable;
            useRateDataGridView.DataSource = useRateTable;
            shiftDataGridView.DataSource = shiftTable;
            operatingDaysDataGridView.DataSource = operatingDaysTable;
            workersDataGridView.DataSource = workerTable;
            sitesDataGridView.DataSource = siteTable;
            ppeDataGridView.DataSource = ppeTable;
            productionRateDataGridView.DataSource = productionRateTable;
            dataValueDataGridView.DataSource = parameterTable;
            remainingValuesDataGridView.DataSource = remainingDataTable;
        }

        void SetColumnWidths()
        {

            // format data grid views
            for (int i = 0; i < dataValueDataGridView.Columns.Count; i++)
            {
                dataValueDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < remainingValuesDataGridView.Columns.Count; i++)
            {
                remainingValuesDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < occupationalExposureDataGridView.Columns.Count; i++)
            {
                occupationalExposureDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < processDescriptionDataGridView.Columns.Count; i++)
            {
                processDescriptionDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < useRateDataGridView.Columns.Count; i++)
            {
                useRateDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < environmentalReleaseDataGridView.Columns.Count; i++)
            {
                environmentalReleaseDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < controlTechnologyDataGridView.Columns.Count; i++)
            {
                controlTechnologyDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < workersDataGridView.Columns.Count; i++)
            {
                workersDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < operatingDaysDataGridView.Columns.Count; i++)
            {
                operatingDaysDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < shiftDataGridView.Columns.Count; i++)
            {
                shiftDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < sitesDataGridView.Columns.Count; i++)
            {
                sitesDataGridView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        void AddTablesToSet()
        {
            genScenarios.Tables.Add(genericScenarioTable);
            genScenarios.Tables.Add(ActivityTable);
            genScenarios.Tables.Add(procDescriptionTable);
            //genScenarios.Tables.Add(infoTable);
            //genScenarios.Tables.Add(occExpTable);
            //genScenarios.Tables.Add(occExposureSummaryTable);
            //genScenarios.Tables.Add(envReleaseTable);
            //genScenarios.Tables.Add(releaseSummaryTable);
            //genScenarios.Tables.Add(productionRateTable);
            genScenarios.Tables.Add(contolTechTable);
            genScenarios.Tables.Add(dataValuesTable);
            //genScenarios.Tables.Add(calculationTable);
            //genScenarios.Tables.Add(concentrationTable);
            //genScenarios.Tables.Add(siteTable);
            //genScenarios.Tables.Add(operatingDaysTable);
            //genScenarios.Tables.Add(workerTable);
            //genScenarios.Tables.Add(shiftTable);
            //genScenarios.Tables.Add(ppeTable);
            //genScenarios.Tables.Add(useRateTable);
            //genScenarios.Tables.Add(parameterTable);
            //genScenarios.Tables.Add(remainingDataTable);
            //genScenarios.Tables.Add(sourceTable);

            releases.Tables.Add(releaseTable);
            releases.Tables.Add(lossFractionReleaseTable);
            releases.Tables.Add(throughputReleaseTable);
            releases.Tables.Add(ap42ReleaseTable);
            releases.Tables.Add(pmnReleaseTable);
            releases.Tables.Add(opptReleaseTable);
            releases.Tables.Add(asssumedReleaseTable);
            releases.Tables.Add(calculatedReleaseTable);
            releases.Tables.Add(agencyReleaseTable);
            releases.Tables.Add(industryReleaseTable);

            exposureData.Tables.Add(exposureTable);
            exposureData.Tables.Add(nioshOshaTable);
            exposureData.Tables.Add(pmnExposureTable);
            exposureData.Tables.Add(opptExposureTable);
            exposureData.Tables.Add(asssumedExpsoureTable);
            exposureData.Tables.Add(calculatedExposureTable);
            exposureData.Tables.Add(agencyExpsoureTable);
            exposureData.Tables.Add(industryExpsoureTable);

            releaseActivities.Tables.Add(cleaningReleaseTable);
            releaseActivities.Tables.Add(dumpingReleaseTable);
            releaseActivities.Tables.Add(dryingReleaseTable);
            releaseActivities.Tables.Add(evaporatingReleaseTable);
            releaseActivities.Tables.Add(fugitiveReleaseTable);
            releaseActivities.Tables.Add(disposalReleaseTable);
            releaseActivities.Tables.Add(residualReleaseTable);
            releaseActivities.Tables.Add(particulateReleaseTable);
            releaseActivities.Tables.Add(samplingReleaseTable);
            releaseActivities.Tables.Add(loadingReleaseTable);
            releaseActivities.Tables.Add(spentReleaseTable);
            releaseActivities.Tables.Add(processReleaseTable);
            releaseActivities.Tables.Add(releaseNotCategorizedTable);

            exposureActivities.Tables.Add(cleaningExposureTable);
            exposureActivities.Tables.Add(dumpingExposureTable);
            exposureActivities.Tables.Add(dryingExposureTable);
            exposureActivities.Tables.Add(evaporatingExposureTable);
            exposureActivities.Tables.Add(fugitiveExposureTable);
            exposureActivities.Tables.Add(disposalExposureTable);
            exposureActivities.Tables.Add(residualExposureTable);
            exposureActivities.Tables.Add(particulateExposureTable);
            exposureActivities.Tables.Add(samplingExposureTable);
            exposureActivities.Tables.Add(loadingExposureTable);
            exposureActivities.Tables.Add(spentExposureTable);
            exposureActivities.Tables.Add(processExposureTable);
            exposureActivities.Tables.Add(expsoureNotCategorizedTable);
        }

        void CleanUpGSNames()
        {
            List<string> scenariosInTable = new List<string>();
            foreach (DataRow r in elementsAndDataTable.Rows)
            {
                if (!scenariosInTable.Contains(r[scenarioElements[1]].ToString()))
                {
                    scenariosInTable.Add(r[scenarioElements[1]].ToString());
                }
            }

            List<string> noMatchingGS = new List<string>();
            foreach (GenericScenario gs in scenarios)
            {
                bool inList = false;
                foreach (string s in scenariosInTable)
                {
                    if (gs.ESD_GS_Name == s)
                    {
                        inList = true;
                        break;
                    }
                }
                if (!inList) noMatchingGS.Add(gs.ESD_GS_Name);
            }

            List<string> gsNotInTable = new List<string>();
            foreach (string s in scenariosInTable)
            {
                bool inList = false;
                foreach (GenericScenario gs in scenarios)
                {
                    if (gs.ESD_GS_Name == s)
                    {
                        inList = true;
                        break;
                    }
                }
                if (!inList) gsNotInTable.Add(s);
            }

            string[] mismatched = new string[]{
                "Enhanced Oil Recovery",
                "Automotive Brake Pad",
                "Biotechnology Premanufacture Notices",
                "Chemical Additives Used in Min",
                "Electrodeposition",
                "Electroplating for Metal Treatment",
                "Dust Releases",
                "Fabric Finishing",
                "Film Deposition",
                "Filtration and Drying Unit Operations",
                "Flexographic Printing",
                "Photoresists",
                "Waterborne Coatings",
                "Granular Detergents",
                "Flexible",
                "Rigid",
                "Printed Circuit Boards",
                "Newspaper",
                "Crude Separation Processes",
                "Metal Products and Machinery",
                "Roll Coating and Curtain Coating",
                "Furniture Industry"
            };

            Dictionary<string, string> gsNameDictionary = new Dictionary<string, string>();
            foreach (string s in mismatched)
            {
                string gsName = string.Empty;
                string tableName = string.Empty;
                foreach (string t in gsNotInTable)
                {
                    if (t.Contains(s))
                    {
                        gsName = t;
                    }
                }
                foreach (string t in noMatchingGS)
                {
                    if (t.Contains(s))
                    {
                        tableName = t;
                    }
                }
                gsNameDictionary.Add(gsName, tableName);
            }
            foreach (DataRow r in elementsAndDataTable.Rows)
            {
                if (gsNameDictionary.ContainsKey(r[1].ToString()))
                {
                    r[1] = gsNameDictionary[r[1].ToString()];
                }
            }
        }

        void CreateElements()
        {
            foreach (DataRow row in elementsAndDataTable.Rows)
            {
                dataElements.Add(new DataElement()
                {
                    Element = row[0].ToString(),
                    ESD_GS_Name = row[1].ToString(),
                    ElementName = row[2].ToString(),
                    Type = row[3].ToString(),
                    Type2 = row[4].ToString(),
                    ExposureType = row[5].ToString(),
                    Activity_Source = row[6].ToString(),
                    mediaOfRelease = row[7].ToString(),
                    SourceSummary = row[8].ToString(),
                    source1 = row[9].ToString(),
                    source2 = row[10].ToString(),
                    source3 = row[11].ToString(),
                    source4 = row[12].ToString(),
                    source5 = row[13].ToString(),
                    source6 = row[14].ToString(),
                    source7 = row[15].ToString(),
                    source8 = row[16].ToString(),
                    Reviewed = row[17].ToString(),
                    ReferenceCheck = row[18].ToString()
                });
            }
        }

        public GenericScenario[] ProcessScenarios()
        {
            List<GenericScenario> retVal = new List<GenericScenario>();
            foreach (DataRow r in infoTable.Rows)
            {
                retVal.Add(new GenericScenario()
                {
                    Category = r[infoColumns[0]].ToString(),
                    DocumentType = r[infoColumns[1]].ToString(),
                    DatePrepared = r[infoColumns[2]].ToString(),
                    ESD_GS_Name = r[infoColumns[3]].ToString(),
                    FullCitation = r[infoColumns[4]].ToString(),
                    DevelopedBy = r[infoColumns[5]].ToString(),
                    Description = r[infoColumns[6]].ToString(),
                    InPaperIndustryDescriptor = r[infoColumns[7]].ToString(),
                    IndustryCodeOrDescription = r[infoColumns[8]].ToString(),
                    IndustryCodeType = r[infoColumns[9]].ToString()
                }
                );
            }
            return retVal.ToArray<GenericScenario>();
        }

        GenericScenario GetScenario(string scenario)
        {
            if (scenario == "GS for Application of Chemicals in Enhanced Oil Recovery") scenario = "GS for Application of Chemicals in Enhanced Oil Recovery: Steam Stimulation, Steam Flooding, and Polymer/Surfactant Flooding";
            if (scenario == "Biotechnology Premanufacture Notices GS") scenario = "GS for Biotechnology Premanufacture Notices";
            if (scenario == "Automotive Brake Pad Replacement") scenario = "Automotive Brake Pad Replacement-GS for Estimating Occupational Exposures";
            if (scenario == "Chemical Additives Used in Min") scenario = "GS for Chemical Additives Used in Mineral and Metal Ore Flotation";
            if (scenario == "Electrodeposition") scenario = "GS for Electrodeposition";
            if (scenario == "Electroplating for Metal Treatment") scenario = "GS for Electroplating for Metal Treatment";
            if (scenario == "Fabric Finishing") scenario = "GS for Fabric Finishing";
            if (scenario == "Film Deposition in IC Fabrications") scenario = "GS for Film Deposition in Integrated Circuit Fabrication";
            if (scenario == "Filtration and Drying Unit Operations") scenario = "GS for Filtration and Drying Unit Operations";
            if (scenario == "Flexographic Printing") scenario = "GS for Flexographic Printing";
            if (scenario == "Formulation of Photoresists") scenario = "GS for Formulation of Photoresists Draft";
            if (scenario == "Granular Detergents Manufacture") scenario = "GS for Granular Detergents Manufacture";
            if (scenario == "Industry Profile for the Flexible PU Foam GS") scenario = "Industry Profile for the Flexible Polyurethane Foam Inudstry - GS";
            if (scenario == "Industry Profile for the Rigid PU Foam GS") scenario = "Industry Profile for the Rigid Polyurethane Foam Inudstry - GS";
            if (scenario == "Material Fabrication Process for Manufacture of Printed Circuit Boards GS") scenario = "Material Fabrication Process for Manufacture of Printed Circuit Boards";
            if (scenario == "Newspaper  Printing GS") scenario = "Newspaper Printing and Cleaning Solvent Use GS";
            if (scenario == "FCC and Crude Separation Processes GS") scenario = "Petroleum Refining Processing Crude Separation Processes and Catalytic Cracking GS";
            if (scenario == "Spray Coatings in the Furniture Industry GS") scenario = "Spray Coatings in the Furniture Industry";
            if (scenario == "Formulation of Waterborne Coatings") scenario = "GS for Formulation of Waterborne Coatings";
            if (scenario == "Metal Products and Machinery Draft GS") scenario = "Metal Products and Machinery";
            if (scenario == "Estimate Dust Releases") scenario = "GS for Estimate Dust Releases from Transfer/Unloading/Loading Operations of Solid Powders";
            if (scenario == "Releases from Roll Coating and Curtain Coating Operations GS") scenario = "Releases from Roll Coating and Curtain Coating Operations";
            foreach (GenericScenario gs in scenarios)
            {
                if (gs.ESD_GS_Name == scenario) return gs;
            }
            return null;
        }

        Source[] GetSources(DataElement el)
        {
            List<Source> sources = new List<Source>();
            if (!string.IsNullOrEmpty(el.source1)) sources.Add(this.GetSource(el.source1));
            if (!string.IsNullOrEmpty(el.source2)) sources.Add(this.GetSource(el.source2));
            if (!string.IsNullOrEmpty(el.source3)) sources.Add(this.GetSource(el.source3));
            if (!string.IsNullOrEmpty(el.source4)) sources.Add(this.GetSource(el.source4));
            if (!string.IsNullOrEmpty(el.source5)) sources.Add(this.GetSource(el.source5));
            if (!string.IsNullOrEmpty(el.source6)) sources.Add(this.GetSource(el.source6));
            if (!string.IsNullOrEmpty(el.source7)) sources.Add(this.GetSource(el.source7));
            if (!string.IsNullOrEmpty(el.source8)) sources.Add(this.GetSource(el.source8));
            return sources.ToArray<Source>();
        }

        Source GetSource(string source)
        {
            foreach (Source s in sources)
            {
                if (s.ReferenceText == source) return s;
            }
            return null;
        }


        void NullsToString()
        {
            foreach (DataRow r in elementsAndDataTable.Rows)
            {
                if (r[scenarioElements[3]] == DBNull.Value) r[scenarioElements[3]] = string.Empty;
                if (r[scenarioElements[4]] == DBNull.Value) r[scenarioElements[4]] = string.Empty;
                if (r[scenarioElements[5]] == DBNull.Value) r[scenarioElements[5]] = string.Empty;
                if (r[scenarioElements[6]] == DBNull.Value) r[scenarioElements[6]] = string.Empty;
                if (r[scenarioElements[7]] == DBNull.Value) r[scenarioElements[7]] = string.Empty;
                if (r[scenarioElements[8]] == DBNull.Value) r[scenarioElements[8]] = string.Empty;
                if (r[scenarioElements[9]] == DBNull.Value) r[scenarioElements[9]] = string.Empty;
                if (r[scenarioElements[10]] == DBNull.Value) r[scenarioElements[10]] = string.Empty;
                if (r[scenarioElements[11]] == DBNull.Value) r[scenarioElements[11]] = string.Empty;
                if (r[scenarioElements[12]] == DBNull.Value) r[scenarioElements[12]] = string.Empty;
                if (r[scenarioElements[13]] == DBNull.Value) r[scenarioElements[13]] = string.Empty;
                if (r[scenarioElements[14]] == DBNull.Value) r[scenarioElements[14]] = string.Empty;
                if (r[scenarioElements[15]] == DBNull.Value) r[scenarioElements[15]] = string.Empty;
                if (r[scenarioElements[16]] == DBNull.Value) r[scenarioElements[16]] = string.Empty;
                if (r[scenarioElements[17]] == DBNull.Value) r[scenarioElements[17]] = string.Empty;
                if (r[scenarioElements[18]] == DBNull.Value) r[scenarioElements[18]] = string.Empty;
            }
        }

        void ExtractSources()
        {
            foreach (DataRow row in elementsAndDataTable.Rows)
            {
                string name = row[scenarioElements[1]].ToString();
                for (int i = 9; i < 17; i++)
                    if (!string.IsNullOrEmpty(row[scenarioElements[i]].ToString()))
                    {
                        bool contained = false;
                        foreach (Source s in sources)
                        {
                            if (s.ReferenceText == row[scenarioElements[i]].ToString()) contained = true;
                        }
                        if (!contained)
                        {
                            sources.Add(new Source
                            {
                                ScenarioName = name,
                                ReferenceText = row[scenarioElements[i]].ToString()
                            });
                        }
                    }
            }

            this.sourceTable.Columns.Add("Generic Scenario");
            this.sourceTable.Columns.Add("Reference");
            foreach (Source s in sources)
            {
                this.sourceTable.Rows.Add(new string[] { s.ScenarioName, s.ReferenceText });
            }
        }

        //string[] ExtractDataRow(DataRow row)
        //{
        //    List<string> retVal = new List<string>();
        //    foreach (string s in scenarioElements)
        //    {
        //        retVal.Add(row[s].ToString());
        //    }
        //    return retVal.ToArray<string>();
        //}

        void ProcessExcel()
        {
            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadSheetDocument = System.IO.File.Exists(@"..\..\Revised Data Element Comparison Draft_2.19.2020_To EPA_with review notes.xlsx") ? DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(@"..\..\Revised Data Element Comparison Draft_2.19.2020_To EPA_with review notes.xlsx", false) : DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(@"Revised Data Element Comparison Draft_2.19.2020_To EPA_with review notes.xlsx", false))
            {
                // DataElementsTable
                DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
                System.Collections.Generic.List<string> sheetNames = new System.Collections.Generic.List<string>();
                foreach (DocumentFormat.OpenXml.Spreadsheet.Sheet s in sheets) { sheetNames.Add(s.Id.Value); }
                string relationshipId = sheets.First().Id.Value;
                DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(sheetNames[0]);
                DocumentFormat.OpenXml.Spreadsheet.Worksheet workSheet = worksheetPart.Worksheet;
                DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = workSheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();

                List<string> columns = new List<string>();
                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(0))
                {
                    elementsAndDataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                    columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = elementsAndDataTable.NewRow();

                    for (int i = 0; i < row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Count(); i++)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(i);
                        tempRow[CellRefToColumn(cell.CellReference)] = GetCellValue(spreadSheetDocument, cell);
                    }

                    elementsAndDataTable.Rows.Add(tempRow);
                }
                elementsAndDataTable.Rows.RemoveAt(0);

                // infoTable
                relationshipId = sheets.First().Id.Value;
                worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(sheetNames[1]);
                workSheet = worksheetPart.Worksheet;
                sheetData = workSheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                rows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();

                columns.Clear();
                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(0))
                {
                    infoTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                    columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = infoTable.NewRow();

                    for (int i = 0; i < row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Count(); i++)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(i);
                        tempRow[CellRefToColumn(cell.CellReference)] = GetCellValue(spreadSheetDocument, cell);
                    }

                    infoTable.Rows.Add(tempRow);
                }

            }
            infoTable.Rows.RemoveAt(0); //...so i'm taking it out here.
        }

        public static int CellRefToColumn(string cellRef)
        {
            if (cellRef.StartsWith("A")) return 0;
            if (cellRef.StartsWith("B")) return 1;
            if (cellRef.StartsWith("C")) return 2;
            if (cellRef.StartsWith("D")) return 3;
            if (cellRef.StartsWith("E")) return 4;
            if (cellRef.StartsWith("F")) return 5;
            if (cellRef.StartsWith("G")) return 6;
            if (cellRef.StartsWith("H")) return 7;
            if (cellRef.StartsWith("I")) return 8;
            if (cellRef.StartsWith("J")) return 9;
            if (cellRef.StartsWith("K")) return 10;
            if (cellRef.StartsWith("L")) return 11;
            if (cellRef.StartsWith("M")) return 12;
            if (cellRef.StartsWith("N")) return 13;
            if (cellRef.StartsWith("O")) return 14;
            if (cellRef.StartsWith("P")) return 15;
            if (cellRef.StartsWith("Q")) return 16;
            if (cellRef.StartsWith("R")) return 17;
            if (cellRef.StartsWith("S")) return 18;
            if (cellRef.StartsWith("T")) return 19;
            return 20;
        }
        public static string GetCellValue(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument document, DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            DocumentFormat.OpenXml.Packaging.SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null) return string.Empty;
            string value = cell.CellValue.InnerXml.Trim();

            if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText.Trim();
            }
            else
            {
                return value;
            }
        }

        void ProcessGenericScenarios(GenericScenario[] gs)
        {
            genericScenarioTable.Columns.Add("Name", typeof(string));
            genericScenarioTable.Columns.Add("Reviewer Name", typeof(string));
            genericScenarioTable.Columns.Add("Description", typeof(string));
            genericScenarioTable.Columns.Add("# of Activities", typeof(string));
            genericScenarioTable.Columns.Add("# of Environmental Releases", typeof(string));
            genericScenarioTable.Columns.Add("# of Occupational Exposures", typeof(string));
            genericScenarioTable.Columns.Add("Facility-level Throughput involving chemical of interest", typeof(string));
            genericScenarioTable.Columns.Add("Concentration of chemical of interest in formulation(s) and/or product(s)", typeof(string));
            genericScenarioTable.Columns.Add("Batch size (if applicable)", typeof(string));
            genericScenarioTable.Columns.Add("Batch duration", typeof(string));
            genericScenarioTable.Columns.Add("Production Rate", typeof(string));
            genericScenarioTable.Columns.Add("Batches per day", typeof(string));
            genericScenarioTable.Columns.Add("Days of operation", typeof(string));
            genericScenarioTable.Columns.Add("Industry Code Type", typeof(string));
            genericScenarioTable.Columns.Add("Industry Code", typeof(string));
            genericScenarioTable.Columns.Add("Facility size distribution", typeof(string));
            genericScenarioTable.Columns.Add("Market share distribution", typeof(string));

            ActivityTable.Columns.Add("Element Number", typeof(string));
            ActivityTable.Columns.Add("Activity", typeof(string));
            ActivityTable.Columns.Add("ChemSTEER Activity", typeof(string));
            ActivityTable.Columns.Add("Generic Scenario Name", typeof(string));
            ActivityTable.Columns.Add("Reviewer Name", typeof(string));
            ActivityTable.Columns.Add("Description", typeof(string));
            ActivityTable.Columns.Add("Type (Release or Occupational Expsoure)", typeof(string));
            ActivityTable.Columns.Add("Pathway or Media", typeof(string));
            ActivityTable.Columns.Add("Numerical Value", typeof(string));
            ActivityTable.Columns.Add("Units of Measure", typeof(string));
            ActivityTable.Columns.Add("Modeled or Measured", typeof(string));
            ActivityTable.Columns.Add("Model or Measurement Description", typeof(string));
            ActivityTable.Columns.Add("Reference Count", typeof(string));

            this.dataValuesTable.Columns.Add("Element Number", typeof(string));
            this.dataValuesTable.Columns.Add("Activity Source", typeof(string));
            this.dataValuesTable.Columns.Add("ChemSTEER Activity", typeof(string));
            this.dataValuesTable.Columns.Add("Generic Scenario Name", typeof(string));
            this.dataValuesTable.Columns.Add("Reviewer Name", typeof(string));
            this.dataValuesTable.Columns.Add("Exposure Pathway (Inhalation, dermal, ingestion)", typeof(string));
            this.dataValuesTable.Columns.Add("Applicable Model Component (Source, Receptor, Compartment, Logistics)", typeof(string));
            this.dataValuesTable.Columns.Add("Model Input Type (Release or Characteristic)", typeof(string));
            this.dataValuesTable.Columns.Add("If Source, Release Compartment? (Facility, Air, Water, Soil)", typeof(string));
            this.dataValuesTable.Columns.Add("Input Variable Name", typeof(string));
            this.dataValuesTable.Columns.Add("Input Variable Quantity", typeof(string));
            this.dataValuesTable.Columns.Add("Input Variable Units", typeof(string));
            this.dataValuesTable.Columns.Add("Input Variable Measured or Estimated?", typeof(string));
            this.dataValuesTable.Columns.Add("If Measured, Data Source Reference?", typeof(string));
            this.dataValuesTable.Columns.Add("If Estimated, Basis? (calculation, engineering judgment, assumption)", typeof(string));
            this.dataValuesTable.Columns.Add("If Calculation, equation used?", typeof(string));
            this.dataValuesTable.Columns.Add("Equation Reference", typeof(string));


            foreach (GenericScenario g in gs)
            {
                DataRow row = genericScenarioTable.NewRow();
                row["Name"] = g.ESD_GS_Name;
                row["Reviewer Name"] = string.Empty;
                row["Description"] = g.InPaperIndustryDescriptor.ToString();
                row["# of Activities"] = g.Activities.Count.ToString();
                row["# of Environmental Releases"] = g.EnvironmentalReleases.Count.ToString();
                row["# of Occupational Exposures"] = g.EnvironmentalReleases.Count.ToString();
                row["Facility-level Throughput involving chemical of interest"] = string.Empty;
                row["Concentration of chemical of interest in formulation(s) and/or product(s)"] = (g.Concentrations.Count == 0) ? "No" : "Yes";
                row["Batch size (if applicable)"] = string.Empty;
                row["Batch duration"] = string.Empty;
                row["Batches per day"] = string.Empty;
                row["Production Rate"] = (g.ProductionRates.Count == 0) ? "No" : "Yes";
                row["Days of operation"] = (g.OperatingDays.Count == 0) ? "No" : "Yes";
                row["Industry Code Type"] = g.IndustryCodeType;
                row["Industry Code"] = string.Join(",", g.IndustryCodes);
                row["Facility size distribution"] = string.Empty;
                row["Market share distribution"] = string.Empty;
                genericScenarioTable.Rows.Add(row);

                foreach (Activity a in g.Activities)
                {
                    foreach (OccupationalExposure oe in a.OccupationalExposures)
                    {
                        DataRow dr = ActivityTable.NewRow();
                        dr["Element Number"] = oe.ElementNumber;
                        dr["Activity"] = a.Name;
                        dr["ChemSTEER Activity"] = a.ChemSTEERActivity;
                        dr["Generic Scenario Name"] = g.ESD_GS_Name;
                        dr["Reviewer Name"] = string.Empty;
                        dr["Description"] = (a.ProcessDescriptions.Count > 0) ? a.ProcessDescriptions[0].Description : string.Empty; ;
                        dr["Type (Release or Occupational Expsoure)"] = "Occupational Exposure";
                        dr["Pathway or Media"] = oe.ExposureType;
                        dr["Numerical Value"] = string.Empty;
                        dr["Units of Measure"] = string.Empty;
                        dr["Modeled or Measured"] = string.Empty;
                        dr["Model or Measurement Description"] = oe.SourceSummary;
                        dr["Reference Count"] = oe.sources.Length.ToString();
                        ActivityTable.Rows.Add(dr);
                    }
                    foreach (EnvironmentalRelease er in a.EnvironmentalReleases)
                    {
                        DataRow dr = ActivityTable.NewRow();
                        dr["Element Number"] = er.ElementNumber;
                        dr["Activity"] = a.Name;
                        dr["ChemSTEER Activity"] = a.ChemSTEERActivity;
                        dr["Generic Scenario Name"] = g.ESD_GS_Name;
                        dr["Reviewer Name"] = string.Empty;
                        dr["Description"] = string.Empty; ;
                        dr["Type (Release or Occupational Expsoure)"] = "Environmental Release";
                        dr["Pathway or Media"] = er.MediaOfRelease;
                        dr["Numerical Value"] = string.Empty;
                        dr["Units of Measure"] = string.Empty;
                        dr["Modeled or Measured"] = string.Empty;
                        dr["Model or Measurement Description"] = er.SourceSummary;
                        dr["Reference Count"] = er.sources.Length.ToString();
                        ActivityTable.Rows.Add(dr);
                    }
                    //foreach (ControlTechnology ct in a.ControlTechnologies)
                    //{
                    //    //    ElementNumber = Int32.Parse(el.Element),
                    //    //ScenarioName = el.ElementName,
                    //    //ElementName = el.ElementName,
                    //    //Type = el.Type,
                    //    //Type2 = el.Type2,
                    //    //SourceSummary = el.SourceSummary
                    //    DataRow dr = ActivityTable.NewRow();
                    //    dr["Element Number"] = ct.ElementNumber;
                    //    dr["Activity"] = a.Name;
                    //    dr["ChemSTEER Activity"] = a.ChemSTEERActivity;
                    //    dr["Generic Scenario Name"] = g.ESD_GS_Name;
                    //    dr["Reviewer Name"] = string.Empty;
                    //    dr["Description"] = string.Empty; ;
                    //    dr["Type (Release or Occupational Expsoure)"] = "Control Technology";
                    //    dr["Pathway or Media"] = string.Empty;
                    //    dr["Numerical Value"] = string.Empty;
                    //    dr["Units of Measure"] = string.Empty;
                    //    dr["Modeled or Measured"] = string.Empty;
                    //    dr["Model or Measurement Description"] = ct.SourceSummary;
                    //    dr["Reference Count"] = ct.sources.Length.ToString();
                    //    ActivityTable.Rows.Add(dr);
                    //}
                }

                foreach (IDataValue dv in g.DataValues)
                {
                    DataRow dr = dataValuesTable.NewRow();
                    dr["Element Number"] = dv.ElementNumber;
                    dr["Activity Source"] = string.Empty;
                    dr["ChemSTEER Activity"] = string.Empty;
                    dr["Generic Scenario Name"] = g.ESD_GS_Name;
                    dr["Reviewer Name"] = string.Empty;
                    dr["Exposure Pathway (Inhalation, dermal, ingestion)"] = string.Empty;
                    dr["Applicable Model Component (Source, Receptor, Compartment, Logistics)"] = string.Empty;
                    dr["Model Input Type (Release or Characteristic)"] = string.Empty;
                    dr["If Source, Release Compartment? (Facility, Air, Water, Soil)"] = string.Empty;
                    dr["Input Variable Name"] = dv.ElementName;
                    dr["Input Variable Quantity"] = string.Empty;
                    dr["Input Variable Units"] = string.Empty;
                    dr["Input Variable Measured or Estimated?"] = dv.SourceSummary;
                    dr["If Measured, Data Source Reference?"] = string.Empty;
                    dr["If Estimated, Basis? (calculation, engineering judgment, assumption)"] = string.Empty;
                    dr["If Calculation, equation used?"] = string.Empty;
                    dr["Equation Reference"] = (dv.Sources.Length > 0) ? dv.Sources[0].ReferenceText : string.Empty;
                    dataValuesTable.Rows.Add(dr);
                }
            }
        }

        private void ExportDataSet(DataSet ds, string destination)
        {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
                {
                    Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets()
                };

                var sheetPart = workbook.WorkbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet;
                DocumentFormat.OpenXml.Spreadsheet.Row headerRow;
                List<String> columns;
                columns = new List<string>();

                foreach (System.Data.DataTable table in ds.Tables)
                {
                    sheetPart = workbook.WorkbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                    sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                    sheetId = 1;
                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                        {
                            DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                        };
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()) //
                            };
                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }
                }
            }
        }
    }
}
