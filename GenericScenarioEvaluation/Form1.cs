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
        DataSet genScenarios = new DataSet();
        DataTable elementsAndDataTable = new DataTable();
        DataTable infoTable = new DataTable("GS Details");
        DataTable occExpTable = new DataTable("Occupational Exposure");
        DataTable concentrationTable = new DataTable("Concntrations");
        DataTable calculationTable = new DataTable("Calculations");
        DataTable procDescriptionTable = new DataTable("Process Descriptions");
        DataTable useRateTable = new DataTable("Use Rates");
        DataTable envReleaseTable = new DataTable("Environmental Releases");
        DataTable contolTechTable = new DataTable("Control Technologies");
        DataTable shiftTable = new DataTable("Shifts");
        DataTable operatingDaysTable = new DataTable("Operating Days");
        DataTable workerTable = new DataTable("Workers");
        DataTable siteTable = new DataTable("Number of Sites");
        DataTable ppeTable = new DataTable("PPE");
        DataTable productionRateTable = new DataTable("ProductionRate");
        DataTable parameterTable = new DataTable("Parameters");
        DataTable remainingDataTable = new DataTable("Data Values");
        DataTable sourceTable = new DataTable("Sources");
        DataTable releaseSummaryTable = new DataTable("Release Summary");
        DataTable occExposureSummaryTable = new DataTable("Occupational Exposure Summary");

        List<DataElement> dataElements = new List<DataElement>();
        List<Source> sources = new List<Source>();
        ExposureCollection expsoures = new ExposureCollection();
        List<string> occActivities = new List<string>();

        List<Concentration> concentrations = new List<Concentration>();
        List<Calculation> calculations = new List<Calculation>();
        List<ProcessDescription> processDescriptions = new List<ProcessDescription>();
        List<UseRate> useRates = new List<UseRate>();
        List<EnvironmentalRelease> envRelease = new List<EnvironmentalRelease>();
        List<string> envReleaseActivities = new List<string>();
        List<ControlTechnology> controlTech = new List<ControlTechnology>();
        List<string> controlTechActivities = new List<string>();
        List<Shift> shifts = new List<Shift>();
        List<OperatingDay> opDays = new List<OperatingDay>();
        List<Worker> workers = new List<Worker>();
        List<Site> sites = new List<Site>();
        List<PPE> ppes = new List<PPE>();
        List<ProductionRate> productions = new List<ProductionRate>();
        List<DataValue> values = new List<DataValue>();
        List<string> uniqueDataElements = new List<string>();
        List<string> uniqueDataSubElements = new List<string>();
        List<RemainingValue> remainingValues = new List<RemainingValue>();

        ReleaseCollection cleaningReleases = new ReleaseCollection();
        ReleaseCollection dumpingReleases = new ReleaseCollection();
        ReleaseCollection dryingReleases = new ReleaseCollection();
        ReleaseCollection evaporatingReleases = new ReleaseCollection();
        ReleaseCollection fugitiveReleases = new ReleaseCollection();
        ReleaseCollection disposalReleases = new ReleaseCollection();
        ReleaseCollection residualReleases = new ReleaseCollection();
        ReleaseCollection particulateReleases = new ReleaseCollection();
        ReleaseCollection samplingReleases = new ReleaseCollection();
        ReleaseCollection loadingReleases = new ReleaseCollection();
        ReleaseCollection spentReleases = new ReleaseCollection();
        ReleaseCollection processReleases = new ReleaseCollection();
        ReleaseCollection releaseNotCategorized = new ReleaseCollection();
        List<string> rCategorized = new List<string>();

        ExposureCollection cleaningOccExp = new ExposureCollection();
        ExposureCollection dryingOccExp = new ExposureCollection();
        ExposureCollection evaporatingOccExp = new ExposureCollection();
        ExposureCollection dumpingOccExp = new ExposureCollection();
        ExposureCollection fugitiveOccExp = new ExposureCollection();
        ExposureCollection disposalOccExp = new ExposureCollection();
        ExposureCollection residualOccExp = new ExposureCollection();
        ExposureCollection particulateOccExp = new ExposureCollection();
        ExposureCollection samplingOccExp = new ExposureCollection();
        ExposureCollection loadingOccExp = new ExposureCollection();
        ExposureCollection spentOccExp = new ExposureCollection();
        ExposureCollection processOccExp = new ExposureCollection();
        ExposureCollection occExpNotCategorized = new ExposureCollection();


        string[] scenarioElements = new string[]{
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
        string[] infoColumns = new string[]{
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
        GenericScenario[] scenarios;

        public Form1()
        {
            InitializeComponent();
            SetUpDataTables();
            ProcessExcel();
            NullsToString();
            ExtractSources();
            AddTablesToSet();
            scenarios = ProcessScenarios();
            CleanUpGSNames();
            CreateElements();

            int count = 0;

            foreach (GenericScenario s in scenarios)
            {
                TreeNode node = treeView2.Nodes.Add(s.ESD_GS_Name, s.ESD_GS_Name);
                TreeNode gsNode = node.Nodes.Add("GSInfo", "GSInfo");
                gsNode.Nodes.Add("Category: " + s.Category);


            }

            var elements = from myElement in dataElements.AsEnumerable()
                           where (myElement.ElementName.ToLower().Contains("process description") ||
                           myElement.Type.ToLower().Contains("process description") ||
                           myElement.ElementName.ToLower().Contains("process summary") ||
                           myElement.ElementName.ToLower().Contains("characterization"))
                           select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                ProcessDescription de = new ProcessDescription()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                processDescriptions.Add(de);
                de.GenericScenario = GetScenario(el.ESD_GS_Name);
                de.GenericScenario.ProcessDescriptions.Add(de);
                de.sources = GetSources(el);
                el.accessed = true;
                procDescriptionTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
                String test = string.Empty;
                if (de.ElementName.StartsWith("Process Description:")) test = de.ElementName;
                else if (!string.IsNullOrEmpty(de.Type2)) test = "Process Description: " + el.Type + ": " + el.Type2;
                else test = "Process Description: " + el.Type;
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test);
                if (!string.IsNullOrEmpty(de.SourceSummary)) node.Nodes.Add(de.SourceSummary);
                if (de.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in de.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            processDescriptionDataGridView.DataSource = procDescriptionTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.ToLower().Contains("occupational")
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                OccupationalExposure o = new OccupationalExposure()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    Type = el.Type,
                    ExposureType = el.ExposureType,
                    ActivitySource = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    sourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(o.ScenarioName);
                o.GenericScenario.OccupationalExposures.Add(o);
                o.sources = GetSources(el);
                expsoures.Add(o);
                el.accessed = true;
                occExpTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.ExposureType, o.Dermal ? "1" : "0", o.DermalSolid ? "1" : "0", o.DermalLiquid ? "1" : "0", o.Inhalation ? "1" : "0", o.Particulate ? "1" : "0", o.ChemicalOrVapor ? "1" : "0", el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                if (!occActivities.Contains(o.ActivitySource)) occActivities.Add(o.ActivitySource);
                if (o.ActivitySource.ToLower().Contains("cleaning"))
                {
                    cleaningOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("dump"))
                {
                    dumpingOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("dry"))
                {
                    dryingOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("evap"))
                {
                    evaporatingOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("fugit")
                    || (o.ActivitySource.ToLower().Contains("off") && o.ActivitySource.ToLower().Contains("gas"))
                    || o.ActivitySource.ToLower().Contains("emissi")
                    || o.ActivitySource.ToLower().Contains("spill"))
                {
                    fugitiveOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("dispos")
                    || (o.ActivitySource.ToLower().Contains("off") && o.ActivitySource.ToLower().Contains("spec"))
                    || o.ActivitySource.ToLower().Contains("waste"))
                {
                    disposalOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("resid"))
                {
                    residualOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("partic")
                    || o.ActivitySource.ToLower().Contains("dust"))
                {
                    particulateOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("sampl"))
                {
                    samplingOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("loadi")
                    || o.ActivitySource.ToLower().Contains("transf")
                    || o.ActivitySource.ToLower().Contains("handl"))
                {
                    loadingOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (o.ActivitySource.ToLower().Contains("spent"))
                {
                    spentOccExp.Add(o);
                    o.ActivityCategorized = true;
                }
                if (!o.ActivityCategorized && !string.IsNullOrEmpty(o.ActivitySource))
                {
                    processOccExp.Add(o);
                    rCategorized.Add(o.ActivitySource);
                }
                String test = "Occupational Exposure";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                node = node.Nodes.ContainsKey("Activity/Source: " + o.ActivitySource) ? node.Nodes["Activity/Source: " + o.ActivitySource] : node.Nodes.Add("Activity/Source: " + o.ActivitySource);
                node = node.Nodes.Add(o.ExposureType);
                if (!string.IsNullOrEmpty(el.Type)) node.Nodes.Add(el.Type);
                if (o.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in o.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            occupationalExposureDataGridView.DataSource = occExpTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Environmental Release") ||
                       myElement.ElementName.Contains("TRI Releases (lb/yr)") ||
                       myElement.ElementName.Contains("Total Industry Estimated Process Water Discharge Flow"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                EnvironmentalRelease er = new EnvironmentalRelease()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    Type2 = el.Type2,
                    ActivitySource = el.Activity_Source,
                    MediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary,
                };
                er.GenericScenario = GetScenario(el.ESD_GS_Name);
                er.sources = GetSources(el);
                er.GenericScenario.EnvironmentalReleases.Add(er);
                envRelease.Add(er);
                el.accessed = true;
                if (!envReleaseActivities.Contains(er.ActivitySource)) envReleaseActivities.Add(er.ActivitySource);
                if (er.ActivitySource.ToLower().Contains("cleaning"))
                {
                    cleaningReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("dump"))
                {
                    dumpingReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("dry"))
                {
                    dryingReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("evap"))
                {
                    evaporatingReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("fugit")
                    || (er.ActivitySource.ToLower().Contains("off") && er.ActivitySource.ToLower().Contains("gas"))
                    || er.ActivitySource.ToLower().Contains("emissi")
                    || er.ActivitySource.ToLower().Contains("spill"))
                {
                    fugitiveReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("dispos")
                    || (er.ActivitySource.ToLower().Contains("off") && er.ActivitySource.ToLower().Contains("spec"))
                    || er.ActivitySource.ToLower().Contains("waste"))
                {
                    disposalReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("resid"))
                {
                    residualReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("partic")
                    || er.ActivitySource.ToLower().Contains("dust"))
                {
                    particulateReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("sampl"))
                {
                    samplingReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("loadi")
                    || er.ActivitySource.ToLower().Contains("transf")
                    || er.ActivitySource.ToLower().Contains("handl"))
                {
                    loadingReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (er.ActivitySource.ToLower().Contains("spent"))
                {
                    spentReleases.Add(er);
                    er.ActivityCategorized = true;
                }
                if (!er.ActivityCategorized && !string.IsNullOrEmpty(er.ActivitySource))
                {
                    processReleases.Add(er);
                    //                    releaseNotCategorized.Add(er);
                }
                if (!(er.RecycledOrReused || er.ToAir || er.ToLand || er.ToWater))
                {

                }
                envReleaseTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, er.ToAir ? "1" : "0", er.ToLand ? "1" : "0", er.ToWater ? "1" : "0", er.RecycledOrReused ? "1" : "0", er.NotSpecified ? "1" : "0", el.SourceSummary });
                String test = "Environmental Releases";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                node = node.Nodes.ContainsKey(er.ActivitySource) ? node.Nodes[er.ActivitySource] : !string.IsNullOrEmpty(er.ActivitySource) ? node.Nodes.Add(er.ActivitySource, er.ActivitySource) : node;
                node = node.Nodes.ContainsKey(er.MediaOfRelease) ? node.Nodes[er.MediaOfRelease] : !string.IsNullOrEmpty(er.MediaOfRelease) ? node.Nodes.Add(er.MediaOfRelease, er.MediaOfRelease) : node;
                if (!string.IsNullOrEmpty(er.Type)) node.Nodes.Add(er.Type);
                if (!string.IsNullOrEmpty(er.Type2)) node.Nodes.Add(er.Type2);
                if (!string.IsNullOrEmpty(er.SourceSummary)) node.Nodes.Add(er.SourceSummary);
                if (er.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in er.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            environmentalReleaseDataGridView.DataSource = envReleaseTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Control Technologies") ||
                       myElement.ElementName.Contains("Treatment Technology"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                ControlTechnology ct = new ControlTechnology()
                {
                    ElementNumber = el.Element,
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
                el.accessed = true;
                controlTechActivities.Add(ct.SourceSummary);
                string test = "Control Technologies";
                contolTechTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(ct.Type)) node.Nodes.Add(ct.Type);
                if (!string.IsNullOrEmpty(ct.Type2)) node.Nodes.Add(ct.Type2);
                if (!string.IsNullOrEmpty(ct.SourceSummary)) node.Nodes.Add(ct.SourceSummary);
                if (ct.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in ct.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            controlTechnologyDataGridView.DataSource = contolTechTable;

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
                Concentration o = new Concentration()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    Type = el.Type,
                    sourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(o.ScenarioName);
                o.GenericScenario.Concentrations.Add(o);
                o.sources = GetSources(el);
                concentrations.Add(o);
                el.accessed = true;
                concentrationTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });

                String test = "Concentration";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                node = node.Nodes.ContainsKey(o.ElementName) ? node.Nodes[o.ElementName] : node.Nodes.Add(o.ElementName);
                node = node.Nodes.Add(o.sourceSummary);
                if (!string.IsNullOrEmpty(el.Type)) node.Nodes.Add(el.Type);
                if (o.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in o.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            ConcentrationDataGridView.DataSource = concentrationTable;

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
                Calculation o = new Calculation()
                {
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    Type = el.Type,
                    ExposureType = el.ExposureType,
                    Activity_Source = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    sourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.GenericScenario.Calculations.Add(o);
                o.sources = GetSources(el);
                calculations.Add(o);
                el.accessed = true;
                calculationTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.ExposureType, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                String test = "Calculations";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                node = node.Nodes.ContainsKey(o.ElementName) ? node.Nodes[o.ElementName] : node.Nodes.Add(o.ElementName);
                if (!string.IsNullOrEmpty(o.Activity_Source)) node = node.Nodes.Add("Activity/Source: " + o.Activity_Source);
                node = node.Nodes.Add(o.sourceSummary);
                if (!string.IsNullOrEmpty(el.Type)) node.Nodes.Add(el.Type);
                if (string.IsNullOrEmpty(o.ExposureType)) node.Nodes.Add("Exposure Type: " + o.ExposureType);
                if (o.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in o.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            CalculationdataGridView.DataSource = calculationTable;

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
                UseRate ur = new UseRate()
                {
                    ElementNumber = el.Element,
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                useRates.Add(ur);
                ur.GenericScenario = GetScenario(el.ESD_GS_Name);
                ur.GenericScenario.UseRates.Add(ur);
                ur.sources = GetSources(el);
                el.accessed = true;
                useRateTable.Rows.Add(new string[] { ur.ElementNumber, el.ESD_GS_Name, ur.ElementName, ur.Type, ur.SourceSummary });
                String test = "Use Rate";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                node = node.Nodes.Add(ur.ElementName);
                if (!string.IsNullOrEmpty(ur.Type)) node.Nodes.Add(ur.Type);
                if (!string.IsNullOrEmpty(ur.SourceSummary)) node.Nodes.Add(ur.SourceSummary);
                if (ur.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in ur.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            useRateDataGridView.DataSource = useRateTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Shift") ||
                       myElement.Type.Contains("Shift"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                Shift shift = new Shift()
                {
                    ElementNumber = el.Element,
                    ScenarioName = el.ESD_GS_Name,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                shift.GenericScenario = GetScenario(el.ESD_GS_Name);
                shift.GenericScenario.Shifts.Add(shift);
                shift.sources = GetSources(el);
                shifts.Add(shift);
                el.accessed = true;
                shiftTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
                string test = "Shifts";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(shift.Type)) node.Nodes.Add(shift.Type);
                if (!string.IsNullOrEmpty(shift.SourceSummary)) node.Nodes.Add(shift.SourceSummary);
                if (shift.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in shift.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            shiftDataGridView.DataSource = shiftTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("Operating") ||
                       myElement.Type.Contains("Operating"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                OperatingDay day = new OperatingDay()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                day.GenericScenario = GetScenario(el.ESD_GS_Name);
                day.GenericScenario.OperatingDays.Add(day);
                day.sources = GetSources(el);
                el.accessed = true;
                opDays.Add(day);
                operatingDaysTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
                string test = "Operating Days";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(day.Type)) node.Nodes.Add(day.Type);
                if (!string.IsNullOrEmpty(day.SourceSummary)) node.Nodes.Add(day.SourceSummary);
                if (day.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in day.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            operatingDaysDataGridView.DataSource = operatingDaysTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.ToLower().Contains("worker") ||
                       myElement.ElementName.ToLower().Contains("operator"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                Worker worker = new Worker()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                worker.GenericScenario = GetScenario(el.ESD_GS_Name);
                worker.GenericScenario.Workers.Add(worker);
                worker.sources = GetSources(el);
                workers.Add(worker);
                workerTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
                string test = "Workers";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(worker.Type)) node.Nodes.Add(worker.Type);
                if (!string.IsNullOrEmpty(worker.SourceSummary)) node.Nodes.Add(worker.SourceSummary);
                if (worker.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in worker.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            workersDataGridView.DataSource = workerTable;


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
                Site site = new Site()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    ScenarioName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                site.GenericScenario = GetScenario(el.ESD_GS_Name);
                site.sources = GetSources(el);
                site.GenericScenario.Sites.Add(site);
                sites.Add(site);
                siteTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
                string test = "Sites";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(site.Type)) node.Nodes.Add(site.Type);
                if (!string.IsNullOrEmpty(site.SourceSummary)) node.Nodes.Add(site.SourceSummary);
                if (site.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in site.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            sitesDataGridView.DataSource = siteTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where (myElement.ElementName.Contains("PPE") ||
                       myElement.Type.Contains("PPE"))
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                PPE pp = new PPE()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                pp.GenericScenario = GetScenario(el.ESD_GS_Name);
                pp.GenericScenario.PPEs.Add(pp);
                pp.sources = GetSources(el);
                ppes.Add(pp);
                ppeTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
                string test = "PPE";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(pp.Type)) node.Nodes.Add(pp.Type);
                if (!string.IsNullOrEmpty(pp.SourceSummary)) node.Nodes.Add(pp.SourceSummary);
                if (pp.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in pp.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            ppeDataGridView.DataSource = ppeTable;

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
                ProductionRate pp = new ProductionRate()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    ScenarioName = el.ESD_GS_Name,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                pp.GenericScenario = GetScenario(el.ESD_GS_Name);
                pp.GenericScenario.ProductionRates.Add(pp);
                pp.GenericScenario.ProductionRates.Add(pp);
                pp.sources = GetSources(el);
                productions.Add(pp);
                productionRateTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
                string test = "Production Rate";
                TreeNode node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (!string.IsNullOrEmpty(pp.Type)) node.Nodes.Add(pp.Type);
                if (!string.IsNullOrEmpty(pp.SourceSummary)) node.Nodes.Add(pp.SourceSummary);
                if (pp.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in pp.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            productionRateDataGridView.DataSource = productionRateTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where !string.IsNullOrEmpty(myElement.Type)
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                DataValue o = new DataValue()
                {
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    ScenarioName = el.ESD_GS_Name,
                    SourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.GenericScenario.DataValues.Add(o);
                o.sources = GetSources(el);
                values.Add(o);
                el.accessed = true;
                if (!uniqueDataElements.Contains(el.ElementName)) uniqueDataElements.Add(el.ElementName);
                if (!uniqueDataSubElements.Contains(el.Type)) uniqueDataSubElements.Add(el.Type);
                parameterTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Type, el.Type2, el.SourceSummary });
                TreeNode node = null;
                if (treeView1.Nodes.ContainsKey(el.ESD_GS_Name))
                    node = treeView1.Nodes[el.ESD_GS_Name];
                else
                    node = treeView1.Nodes.Add(el.ESD_GS_Name, el.ESD_GS_Name);
                if (node.Nodes.ContainsKey(el.ElementName))
                    node = node.Nodes[el.ElementName];
                else
                    node = node.Nodes.Add(el.ElementName, el.ElementName);
                if (!string.IsNullOrEmpty(el.Type)) node = node.Nodes.Add(el.Type, el.Type);
                if (!string.IsNullOrEmpty(el.Type2)) node = node.Nodes.Add(el.Type2);
                node.Nodes.Add(el.SourceSummary);
                string test = "Data Values";
                node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (node.Nodes.ContainsKey(el.ElementName))
                    node = node.Nodes[el.ElementName];
                else
                    node = node.Nodes.Add(el.ElementName, el.ElementName);
                if (!string.IsNullOrEmpty(el.Type)) node = node.Nodes.Add(el.Type, el.Type);
                if (!string.IsNullOrEmpty(el.Type2)) node = node.Nodes.Add(el.Type2);
                if (o.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in o.sources) node.Nodes.Add(s.ReferenceText);
                }
            }
            dataValueDataGridView.DataSource = parameterTable;

            elements = from myElement in dataElements.AsEnumerable()
                       where !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                RemainingValue o = new RemainingValue()
                {
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
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
                o.sources = GetSources(el);
                remainingValues.Add(o);
                if (!uniqueDataElements.Contains(el.ElementName)) uniqueDataElements.Add(el.ElementName);
                if (!uniqueDataSubElements.Contains(el.Type)) uniqueDataSubElements.Add(el.Type);
                remainingDataTable.Rows.Add(new string[] { el.Element, el.ESD_GS_Name, el.ElementName, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
                TreeNode node = treeView1.Nodes[el.ESD_GS_Name];
                if (node == null)
                    node = treeView1.Nodes.Add(el.ESD_GS_Name, el.ESD_GS_Name);
                if (node.Nodes.ContainsKey(el.ElementName))
                    node = node.Nodes[el.ElementName];
                else
                    node = node.Nodes.Add(el.ElementName, el.ElementName);
                if (!string.IsNullOrEmpty(el.Type)) node.Nodes.Add(el.Type);
                node.Nodes.Add(el.SourceSummary);
                string test = "Data Values";
                node = treeView2.Nodes[el.ESD_GS_Name].Nodes.ContainsKey(test) ? treeView2.Nodes[el.ESD_GS_Name].Nodes[test] : treeView2.Nodes[el.ESD_GS_Name].Nodes.Add(test, test);
                if (node.Nodes.ContainsKey(el.ElementName))
                    node = node.Nodes[el.ElementName];
                else
                    node = node.Nodes.Add(el.ElementName, el.ElementName);
                if (!string.IsNullOrEmpty(el.Type)) node = node.Nodes.Add(el.Type, el.Type);
                if (!string.IsNullOrEmpty(el.Type2)) node = node.Nodes.Add(el.Type2);
                if (o.sources.Length > 0)
                {
                    node = node.Nodes.Add("Sources", "Sources");
                    foreach (Source s in o.sources) node.Nodes.Add(s.ReferenceText);
                }

            }
            remainingValuesDataGridView.DataSource = remainingDataTable;
            SetColumnWidths();
            int numrows = 0;
            foreach (DataTable table in genScenarios.Tables)
            {
                numrows = numrows + table.Rows.Count;
            }

            envReleaseActivities.Sort();
            occActivities.Sort();



            this.releaseSummaryTable.Columns.Add(new DataColumn("Activity"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release to Air"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release to Land"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release to Water"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Release Not Specified"));
            this.releaseSummaryTable.Columns.Add(new DataColumn("Total Releases"));
            this.releaseSummaryTable.Rows.Add(new string[] { "Cleaning", cleaningReleases.ToAir.ToString(), cleaningReleases.ToLand.ToString(), cleaningReleases.ToWater.ToString(), cleaningReleases.NotSpecified.ToString(), cleaningReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Dumping", dumpingReleases.ToAir.ToString(), dumpingReleases.ToLand.ToString(), dumpingReleases.ToWater.ToString(), dumpingReleases.NotSpecified.ToString(), dumpingReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Drying", dryingReleases.ToAir.ToString(), dryingReleases.ToLand.ToString(), dryingReleases.ToWater.ToString(), dryingReleases.NotSpecified.ToString(), dryingReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Evaporating", evaporatingReleases.ToAir.ToString(), evaporatingReleases.ToLand.ToString(), evaporatingReleases.ToWater.ToString(), evaporatingReleases.NotSpecified.ToString(), evaporatingReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Fugitive", fugitiveReleases.ToAir.ToString(), fugitiveReleases.ToLand.ToString(), fugitiveReleases.ToWater.ToString(), fugitiveReleases.NotSpecified.ToString(), fugitiveReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Disposal", disposalReleases.ToAir.ToString(), disposalReleases.ToLand.ToString(), disposalReleases.ToWater.ToString(), disposalReleases.NotSpecified.ToString(), disposalReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Residual", residualReleases.ToAir.ToString(), residualReleases.ToLand.ToString(), residualReleases.ToWater.ToString(), residualReleases.NotSpecified.ToString(), residualReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Particulate", particulateReleases.ToAir.ToString(), particulateReleases.ToLand.ToString(), particulateReleases.ToWater.ToString(), particulateReleases.NotSpecified.ToString(), particulateReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Sampling", samplingReleases.ToAir.ToString(), samplingReleases.ToLand.ToString(), samplingReleases.ToWater.ToString(), samplingReleases.NotSpecified.ToString(), samplingReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Loading", loadingReleases.ToAir.ToString(), loadingReleases.ToLand.ToString(), loadingReleases.ToWater.ToString(), loadingReleases.NotSpecified.ToString(), loadingReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Spent Materials", spentReleases.ToAir.ToString(), spentReleases.ToLand.ToString(), spentReleases.ToWater.ToString(), spentReleases.NotSpecified.ToString(), spentReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Process", processReleases.ToAir.ToString(), processReleases.ToLand.ToString(), processReleases.ToWater.ToString(), processReleases.NotSpecified.ToString(), processReleases.Count.ToString() });
            this.releaseSummaryTable.Rows.Add(new string[] { "Not Specified", releaseNotCategorized.ToAir.ToString(), releaseNotCategorized.ToLand.ToString(), releaseNotCategorized.ToWater.ToString(), releaseNotCategorized.NotSpecified.ToString(), releaseNotCategorized.Count.ToString() });

            this.occExposureSummaryTable.Columns.Add(new DataColumn("Activity"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Chemical Vapor Inhalation"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Particulate Inhalation"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Inhalation Not Specified"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Total Inhalation"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Liquid Dermal"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Solid Dermal"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Dermal Not Categorized"));
            this.occExposureSummaryTable.Columns.Add(new DataColumn("Total Dermal"));
            this.occExposureSummaryTable.Rows.Add(new string[] { "Cleaning", cleaningOccExp.ChemicalOrVapor.ToString(), cleaningOccExp.ParticulateInhalation.ToString(), cleaningOccExp.InhalationNotSpecified.ToString(), cleaningOccExp.TotalInhalation.ToString(),
                cleaningOccExp.DermalLiquid.ToString(), cleaningOccExp.DermalSolid.ToString(), cleaningOccExp.DermalNotCategorized.ToString(), cleaningOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Dumping", dumpingOccExp.ChemicalOrVapor.ToString(), dumpingOccExp.ParticulateInhalation.ToString(), dumpingOccExp.InhalationNotSpecified.ToString(), dumpingOccExp.TotalInhalation.ToString(),
                dumpingOccExp.DermalLiquid.ToString(), dumpingOccExp.DermalSolid.ToString(), dumpingOccExp.DermalNotCategorized.ToString(), dumpingOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Drying", dryingOccExp.ChemicalOrVapor.ToString(), dryingOccExp.ParticulateInhalation.ToString(), dryingOccExp.InhalationNotSpecified.ToString(), dryingOccExp.TotalInhalation.ToString(),
                dryingOccExp.DermalLiquid.ToString(), dryingOccExp.DermalSolid.ToString(), dryingOccExp.DermalNotCategorized.ToString(), dryingOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Evaporating", evaporatingOccExp.ChemicalOrVapor.ToString(), evaporatingOccExp.ParticulateInhalation.ToString(), evaporatingOccExp.InhalationNotSpecified.ToString(), evaporatingOccExp.TotalInhalation.ToString(),
                evaporatingOccExp.DermalLiquid.ToString(), evaporatingOccExp.DermalSolid.ToString(), evaporatingOccExp.DermalNotCategorized.ToString(), evaporatingOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Fugitive", fugitiveOccExp.ChemicalOrVapor.ToString(), fugitiveOccExp.ParticulateInhalation.ToString(), fugitiveOccExp.InhalationNotSpecified.ToString(), fugitiveOccExp.TotalInhalation.ToString(),
                fugitiveOccExp.DermalLiquid.ToString(), fugitiveOccExp.DermalSolid.ToString(), fugitiveOccExp.DermalNotCategorized.ToString(), fugitiveOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Disposal", disposalOccExp.ChemicalOrVapor.ToString(), disposalOccExp.ParticulateInhalation.ToString(), disposalOccExp.InhalationNotSpecified.ToString(), disposalOccExp.TotalInhalation.ToString(),
                disposalOccExp.DermalLiquid.ToString(), disposalOccExp.DermalSolid.ToString(), disposalOccExp.DermalNotCategorized.ToString(), disposalOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Residual", residualOccExp.ChemicalOrVapor.ToString(), residualOccExp.ParticulateInhalation.ToString(), residualOccExp.InhalationNotSpecified.ToString(), residualOccExp.TotalInhalation.ToString(),
                residualOccExp.DermalLiquid.ToString(), residualOccExp.DermalSolid.ToString(), residualOccExp.DermalNotCategorized.ToString(), residualOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Particulate", particulateOccExp.ChemicalOrVapor.ToString(), particulateOccExp.ParticulateInhalation.ToString(), particulateOccExp.InhalationNotSpecified.ToString(), particulateOccExp.TotalInhalation.ToString(),
                particulateOccExp.DermalLiquid.ToString(), particulateOccExp.DermalSolid.ToString(), particulateOccExp.DermalNotCategorized.ToString(), particulateOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Sampling", samplingOccExp.ChemicalOrVapor.ToString(), samplingOccExp.ParticulateInhalation.ToString(), samplingOccExp.InhalationNotSpecified.ToString(), samplingOccExp.TotalInhalation.ToString(),
                samplingOccExp.DermalLiquid.ToString(), samplingOccExp.DermalSolid.ToString(), samplingOccExp.DermalNotCategorized.ToString(), samplingOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Loading", loadingOccExp.ChemicalOrVapor.ToString(), loadingOccExp.ParticulateInhalation.ToString(), loadingOccExp.InhalationNotSpecified.ToString(), loadingOccExp.TotalInhalation.ToString(),
                loadingOccExp.DermalLiquid.ToString(), loadingOccExp.DermalSolid.ToString(), loadingOccExp.DermalNotCategorized.ToString(), loadingOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Spent Materials", spentOccExp.ChemicalOrVapor.ToString(), spentOccExp.ParticulateInhalation.ToString(), spentOccExp.InhalationNotSpecified.ToString(), spentOccExp.TotalInhalation.ToString(),
                spentOccExp.DermalLiquid.ToString(), spentOccExp.DermalSolid.ToString(), spentOccExp.DermalNotCategorized.ToString(), spentOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Process", processOccExp.ChemicalOrVapor.ToString(), processOccExp.ParticulateInhalation.ToString(), processOccExp.InhalationNotSpecified.ToString(), processOccExp.TotalInhalation.ToString(),
                processOccExp.DermalLiquid.ToString(), processOccExp.DermalSolid.ToString(), processOccExp.DermalNotCategorized.ToString(), processOccExp.TotalDermal.ToString()});
            this.occExposureSummaryTable.Rows.Add(new string[] { "Not Specified", occExpNotCategorized.ChemicalOrVapor.ToString(), occExpNotCategorized.ParticulateInhalation.ToString(), occExpNotCategorized.InhalationNotSpecified.ToString(), occExpNotCategorized.TotalInhalation.ToString(),
                occExpNotCategorized.DermalLiquid.ToString(), occExpNotCategorized.DermalSolid.ToString(), occExpNotCategorized.DermalNotCategorized.ToString(), occExpNotCategorized.TotalDermal.ToString()});



            string output = "Activity\tChemical Vapor Inhalation\tParticulate Inhalation\tInhalation NotSpecified\tTotal Inhalation\tLiquid Dermal\tSolid Dermal\tDermal Not Categorized\tTotal Dermal\t";
            output = output + "Release to Air\tReleases to Land\tReleases To Water\tRelease Not Specified\tTotal Releases\n";
            output = output + "Cleaning \t" + cleaningOccExp.ChemicalOrVapor + "\t" + cleaningOccExp.ParticulateInhalation + "\t" + cleaningOccExp.InhalationNotSpecified + "\t" + cleaningOccExp.TotalInhalation
                + "\t" + cleaningOccExp.DermalLiquid + "\t" + cleaningOccExp.DermalSolid + "\t" + cleaningOccExp.DermalNotCategorized + "\t" + cleaningOccExp.TotalDermal + "\t" + cleaningReleases.ToAir
                + "\t" + cleaningReleases.ToLand + "\t" + cleaningReleases.ToWater + "\t" + cleaningReleases.NotSpecified + "\t" + cleaningReleases.Count + "\n";
            output = output + "Dumping \t" + dumpingOccExp.ChemicalOrVapor + "\t" + dumpingOccExp.ParticulateInhalation + "\t" + dumpingOccExp.InhalationNotSpecified + "\t" + +dumpingOccExp.TotalInhalation
                + "\t" + dumpingOccExp.DermalLiquid + "\t" + dumpingOccExp.DermalSolid + "\t" + dumpingOccExp.DermalNotCategorized + "\t" + dumpingOccExp.TotalDermal + "\t" + dumpingReleases.ToAir
                + "\t" + dumpingReleases.ToLand + "\t" + dumpingReleases.ToWater + "\t" + dumpingReleases.NotSpecified + "\t" + dumpingReleases.Count + "\n";
            output = output + "Drying \t" + dryingOccExp.ChemicalOrVapor + "\t" + dryingOccExp.ParticulateInhalation + "\t" + dryingOccExp.InhalationNotSpecified + "\t" + +dryingOccExp.TotalInhalation
                + "\t" + dryingOccExp.DermalLiquid + "\t" + dryingOccExp.DermalSolid + "\t" + dryingOccExp.DermalNotCategorized + "\t" + dryingOccExp.TotalDermal + "\t" + dryingReleases.ToAir
                + "\t" + dryingReleases.ToLand + "\t" + dryingReleases.ToWater + "\t" + dryingReleases.NotSpecified + "\t" + dryingReleases.Count + "\n";
            output = output + "Evaporating \t" + evaporatingOccExp.ChemicalOrVapor + "\t" + evaporatingOccExp.ParticulateInhalation + "\t" + evaporatingOccExp.InhalationNotSpecified + "\t" + evaporatingOccExp.TotalInhalation
                + "\t" + evaporatingOccExp.DermalLiquid + "\t" + evaporatingOccExp.DermalSolid + "\t" + evaporatingOccExp.DermalNotCategorized + "\t" + evaporatingOccExp.TotalDermal + "\t" + evaporatingReleases.ToAir
                + "\t" + evaporatingReleases.ToLand + "\t" + evaporatingReleases.ToWater + "\t" + evaporatingReleases.NotSpecified + "\t" + evaporatingReleases.Count + "\n";
            output = output + "Fugitive \t" + fugitiveOccExp.ChemicalOrVapor + "\t" + fugitiveOccExp.ParticulateInhalation + "\t" + fugitiveOccExp.InhalationNotSpecified + "\t" + fugitiveOccExp.TotalInhalation
                + "\t" + fugitiveOccExp.DermalLiquid + "\t" + fugitiveOccExp.DermalSolid + "\t" + fugitiveOccExp.DermalNotCategorized + "\t" + fugitiveOccExp.TotalDermal + "\t" + fugitiveReleases.ToAir
                + "\t" + fugitiveReleases.ToLand + "\t" + fugitiveReleases.ToWater + "\t" + fugitiveReleases.NotSpecified + "\t" + fugitiveReleases.Count + "\n";
            output = output + "Disposal \t" + disposalOccExp.ChemicalOrVapor + "\t" + disposalOccExp.ParticulateInhalation + "\t" + disposalOccExp.InhalationNotSpecified + "\t" + disposalOccExp.TotalInhalation
                + "\t" + disposalOccExp.DermalLiquid + "\t" + disposalOccExp.DermalSolid + "\t" + disposalOccExp.DermalNotCategorized + "\t" + disposalOccExp.TotalDermal + "\t" + disposalReleases.ToAir
                + "\t" + disposalReleases.ToLand + "\t" + disposalReleases.ToWater + "\t" + disposalReleases.NotSpecified + "\t" + disposalReleases.Count + "\n";
            output = output + "Residual \t" + residualOccExp.ChemicalOrVapor + "\t" + residualOccExp.ParticulateInhalation + "\t" + residualOccExp.InhalationNotSpecified + "\t" + residualOccExp.TotalInhalation
                + "\t" + residualOccExp.DermalLiquid + "\t" + residualOccExp.DermalSolid + "\t" + residualOccExp.DermalNotCategorized + "\t" + residualOccExp.TotalDermal + "\t" + residualReleases.ToAir
                + "\t" + residualReleases.ToLand + "\t" + residualReleases.ToWater + "\t" + residualReleases.NotSpecified + "\t" + residualReleases.Count + "\n";
            output = output + "Particulate \t" + particulateOccExp.ChemicalOrVapor + "\t" + particulateOccExp.ParticulateInhalation + "\t" + particulateOccExp.InhalationNotSpecified + "\t" + particulateOccExp.TotalInhalation
                + "\t" + particulateOccExp.DermalLiquid + "\t" + particulateOccExp.DermalSolid + "\t" + particulateOccExp.DermalNotCategorized + "\t" + particulateOccExp.TotalDermal + "\t" + particulateReleases.ToAir
                + "\t" + particulateReleases.ToLand + "\t" + particulateReleases.ToWater + "\t" + particulateReleases.NotSpecified + "\t" + particulateReleases.Count + "\n";
            output = output + "Sampling \t" + samplingOccExp.ChemicalOrVapor + "\t" + samplingOccExp.ParticulateInhalation + "\t" + samplingOccExp.InhalationNotSpecified + "\t" + samplingOccExp.TotalInhalation
                + "\t" + samplingOccExp.DermalLiquid + "\t" + samplingOccExp.DermalSolid + "\t" + samplingOccExp.DermalNotCategorized + "\t" + samplingOccExp.TotalDermal + "\t" + samplingReleases.ToAir
                + "\t" + samplingReleases.ToLand + "\t" + samplingReleases.ToWater + "\t" + samplingReleases.NotSpecified + "\t" + samplingReleases.Count + "\n";
            output = output + "Loading \t" + loadingOccExp.ChemicalOrVapor + "\t" + loadingOccExp.ParticulateInhalation + "\t" + loadingOccExp.InhalationNotSpecified + "\t" + loadingOccExp.TotalInhalation
                + "\t" + loadingOccExp.DermalLiquid + "\t" + loadingOccExp.DermalSolid + "\t" + loadingOccExp.DermalNotCategorized + "\t" + loadingOccExp.TotalDermal + "\t" + loadingReleases.ToAir
                + "\t" + loadingReleases.ToLand + "\t" + loadingReleases.ToWater + "\t" + loadingReleases.NotSpecified + "\t" + loadingReleases.Count + "\n";
            output = output + "Spent Materials \t" + spentOccExp.ChemicalOrVapor + "\t" + spentOccExp.ParticulateInhalation + "\t" + spentOccExp.InhalationNotSpecified + "\t" + spentOccExp.TotalInhalation
                + "\t" + spentOccExp.DermalLiquid + "\t" + spentOccExp.DermalSolid + "\t" + spentOccExp.DermalNotCategorized + "\t" + spentOccExp.TotalDermal + "\t" + spentReleases.ToAir
                + "\t" + spentReleases.ToLand + "\t" + spentReleases.ToWater + "\t" + spentReleases.NotSpecified + "\t" + spentReleases.Count + "\n";
            output = output + "Process \t" + processOccExp.ChemicalOrVapor + "\t" + processOccExp.ParticulateInhalation + "\t" + processOccExp.InhalationNotSpecified + "\t" + processOccExp.TotalInhalation
                + "\t" + processOccExp.DermalLiquid + "\t" + processOccExp.DermalSolid + "\t" + processOccExp.DermalNotCategorized + "\t" + processOccExp.TotalDermal + "\t" + processReleases.ToAir
                + "\t" + processReleases.ToLand + "\t" + processReleases.ToWater + "\t" + processReleases.NotSpecified + "\t" + cleaningReleases.Count + "\n";
            output = output + "Not Categorized \t" + occExpNotCategorized.ChemicalOrVapor + "\t" + occExpNotCategorized.ParticulateInhalation + "\t" + occExpNotCategorized.InhalationNotSpecified + "\t" + occExpNotCategorized.TotalInhalation
                + "\t" + occExpNotCategorized.DermalLiquid + "\t" + occExpNotCategorized.DermalSolid + "\t" + occExpNotCategorized.DermalNotCategorized + "\t" + occExpNotCategorized.TotalDermal + "\t" + releaseNotCategorized.ToAir + "\t" + releaseNotCategorized.ToLand
                + "\t" + releaseNotCategorized.ToWater + "\t" + releaseNotCategorized.NotSpecified + "\t" + releaseNotCategorized.Count + "\n";
            System.IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\table.txt", output);
            ExportDataSet(scenarios, genScenarios, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\GenericScenarioOutputs.xlsx");
        }

        void SetUpDataTables()
        {
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
            this.procDescriptionTable.Columns.Add("Element Name");
            this.procDescriptionTable.Columns.Add("Element Type");
            this.procDescriptionTable.Columns.Add("Element Type 2");
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
            this.envReleaseTable.Columns.Add("Recycked or Reused");
            this.envReleaseTable.Columns.Add("Not Specified");
            this.envReleaseTable.Columns.Add("Source Summary");

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
            genScenarios.Tables.Add(procDescriptionTable);
            genScenarios.Tables.Add(infoTable);
            genScenarios.Tables.Add(occExpTable);
            genScenarios.Tables.Add(occExposureSummaryTable);
            genScenarios.Tables.Add(envReleaseTable);
            genScenarios.Tables.Add(releaseSummaryTable);
            genScenarios.Tables.Add(productionRateTable);
            genScenarios.Tables.Add(contolTechTable);
            genScenarios.Tables.Add(calculationTable);
            genScenarios.Tables.Add(concentrationTable);
            genScenarios.Tables.Add(siteTable);
            genScenarios.Tables.Add(operatingDaysTable);
            genScenarios.Tables.Add(workerTable);
            genScenarios.Tables.Add(shiftTable);
            genScenarios.Tables.Add(ppeTable);
            genScenarios.Tables.Add(useRateTable);
            genScenarios.Tables.Add(parameterTable);
            genScenarios.Tables.Add(remainingDataTable);
            genScenarios.Tables.Add(sourceTable);
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

        string[] ExtractDataRow(DataRow row)
        {
            List<string> retVal = new List<string>();
            foreach (string s in scenarioElements)
            {
                retVal.Add(row[s].ToString());
            }
            return retVal.ToArray<string>();
        }

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
                        tempRow[cellRefToColumn(cell.CellReference)] = GetCellValue(spreadSheetDocument, cell);
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
                        tempRow[cellRefToColumn(cell.CellReference)] = GetCellValue(spreadSheetDocument, cell);
                    }

                    infoTable.Rows.Add(tempRow);
                }

            }
            infoTable.Rows.RemoveAt(0); //...so i'm taking it out here.
        }

        public static int cellRefToColumn(string cellRef)
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

        private void ExportDataSet(GenericScenario[] gs, DataSet ds, string destination)
        {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

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

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Data Element Types" };
                sheets.Append(sheet);

                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (string column in gs[0].GetColumns())
                {
                    columns.Add(column);
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (GenericScenario g in gs)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        if (Int32.TryParse(g[col], out int temp)) cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(g[col]); //
                        newRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(newRow);
                }


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
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }
                }
            }
        }
    }
}
