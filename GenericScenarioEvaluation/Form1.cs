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
        DataTable infoTable = new DataTable();
        DataTable occExpTable = new DataTable("Occupational Exposure");
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

        List<DataElement> dataElements = new List<DataElement>();
        List<Source> sources = new List<Source>();
        List<OccupationalExposure> expsoures = new List<OccupationalExposure>();
        List<ProcessDescription> processDescriptions = new List<ProcessDescription>();
        List<UseRate> useRates = new List<UseRate>();
        List<EnvironmentalRelease> envRelease = new List<EnvironmentalRelease>();
        List<ControlTechnology> controlTech = new List<ControlTechnology>();
        List<Shift> shifts = new List<Shift>();
        List<OperatingDays> opDays = new List<OperatingDays>();
        List<Worker> workers = new List<Worker>();
        List<Site> sites = new List<Site>();
        List<PPE> ppes = new List<PPE>();
        List<ProductionRate> productions = new List<ProductionRate>();
        List<DataValue> values = new List<DataValue>();
        List<string> uniqueDataElements = new List<string>();
        List<string> uniqueDataSubElements = new List<string>();
        List<RemainingValue> remainingValues = new List<RemainingValue>();

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
        string accessed = "Accessed";
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

            var elements = from myElement in dataElements.AsEnumerable()
                           where myElement.ElementName.Contains("Occupational")
                           && !myElement.ElementName.Contains("Process Description")
                           select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                OccupationalExposure o = new OccupationalExposure()
                {
                    ScenarioName = el.ElementName,
                    ElementNumber = el.Element,
                    Type = el.Type,
                    ExposureType = el.ExposureType,
                    Activity_Source = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    sourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(o.ScenarioName);
                o.sources = GetSources(el);
                expsoures.Add(o);
                el.accessed = true;
                occExpTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.ExposureType, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            }
            occupationalExposureDataGridView.DataSource = occExpTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Process Description")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                ProcessDescription de = new ProcessDescription()
                {
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                processDescriptions.Add(de);
                de.GenericScenario = GetScenario(el.ESD_GS_Name);
                de.sources = GetSources(el);
                el.accessed = true;
                procDescriptionTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }
            processDescriptionDataGridView.DataSource = procDescriptionTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.ToLower().Contains("use rate") ||
                       myElement.Type.ToLower().Contains("use rate") ||
                       myElement.ElementName.ToLower().Contains("daily use") ||
                       myElement.ElementName.ToLower().Contains("annual use")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                UseRate ur = new UseRate()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                useRates.Add(ur);
                ur.GenericScenario = GetScenario(el.ESD_GS_Name);
                ur.sources = GetSources(el);
                el.accessed = true;
                useRateTable.Rows.Add(new string[] { ur.ElementNumber, ur.ElementName, ur.Type, ur.SourceSummary });
            }
            useRateDataGridView.DataSource = useRateTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Environmental Release") ||
                       myElement.ElementName.Contains("TRI Releases (lb/yr)") ||
                       myElement.ElementName.Contains("Total Industry Estimated Process Water Discharge Flow")
                       && !myElement.ElementName.ToLower().Contains("Process Description")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                EnvironmentalRelease er = new EnvironmentalRelease()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    ActivitySource = el.Activity_Source,
                    MediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary
                };
                er.GenericScenario = GetScenario(el.ESD_GS_Name);
                er.sources = GetSources(el);
                envRelease.Add(er);
                el.accessed = true;
                envReleaseTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            }
            environmentalReleaseDataGridView.DataSource = envReleaseTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Control Technologies") ||
                       myElement.ElementName.Contains("Treatment Technology")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                ControlTechnology ct = new ControlTechnology()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                ct.GenericScenario = GetScenario(el.ESD_GS_Name);
                ct.sources = GetSources(el);
                controlTech.Add(ct);
                el.accessed = true;
                contolTechTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }
            controlTechnologyDataGridView.DataSource = contolTechTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Shift") ||
                       myElement.Type.Contains("Shift")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                Shift shift = new Shift()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                shift.GenericScenario = GetScenario(el.ESD_GS_Name);
                shift.sources = GetSources(el);
                shifts.Add(shift);
                el.accessed = true;
                shiftTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }
            shiftDataGridView.DataSource = shiftTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Operating")
                       && !myElement.Type.Contains("Process Description")
                       && !myElement.Type.Contains("Use Rate")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                OperatingDays day = new OperatingDays()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    Type2 = el.Type2,
                    SourceSummary = el.SourceSummary
                };
                day.GenericScenario = GetScenario(el.ESD_GS_Name);
                day.sources = GetSources(el);
                el.accessed = true;
                opDays.Add(day);
                operatingDaysTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            }
            operatingDaysDataGridView.DataSource = operatingDaysTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Worker")
                       && !myElement.Type.Contains("Process Description")
                       && !myElement.Type.Contains("Use Rate")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                Worker worker = new Worker()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                worker.GenericScenario = GetScenario(el.ESD_GS_Name);
                worker.sources = GetSources(el);
                workers.Add(worker);
                workerTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
            }
            workersDataGridView.DataSource = workerTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("Number of Sites") ||
                        myElement.Type.Contains("Number of Sites")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                Site site = new Site()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                site.GenericScenario = GetScenario(el.ESD_GS_Name);
                site.sources = GetSources(el);
                sites.Add(site);
                siteTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
            }
            sitesDataGridView.DataSource = siteTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.Contains("PPE") ||
                        myElement.Type.Contains("PPE")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                PPE pp = new PPE()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                pp.GenericScenario = GetScenario(el.ESD_GS_Name);
                pp.sources = GetSources(el);
                ppes.Add(pp);
                ppeTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
            }
            ppeDataGridView.DataSource = ppeTable;


            elements = from myElement in dataElements.AsEnumerable()
                       where myElement.ElementName.ToLower().Contains("production")
                       && !myElement.Type.Contains("Process Description")
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                ProductionRate pp = new ProductionRate()
                {
                    ElementNumber = el.Element,
                    ElementName = el.ElementName,
                    Type = el.Type,
                    SourceSummary = el.SourceSummary
                };
                pp.GenericScenario = GetScenario(el.ESD_GS_Name);
                pp.sources = GetSources(el);
                productions.Add(pp);
                productionRateTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
                el.accessed = true;
            }
            productionRateDataGridView.DataSource = productionRateTable;

            
            elements = from myElement in dataElements.AsEnumerable()
                       where !string.IsNullOrEmpty(myElement.ElementName)
                       && !string.IsNullOrEmpty(myElement.Type)
                       && !myElement.accessed
                       select myElement;

            foreach (DataElement el in elements)
            {
                count++;
                DataValue o = new DataValue()
                {
                    ElementName = el.ElementName,
                    ElementNumber = el.Element,
                    SourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.sources = GetSources(el);
                values.Add(o);
                if (!uniqueDataElements.Contains(el.ElementName)) uniqueDataElements.Add(el.ElementName);
                if (!uniqueDataSubElements.Contains(el.Type)) uniqueDataSubElements.Add(el.Type);
                parameterTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
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
                    Type = el.Type,
                    Type2 = el.Type2,
                    ExposureType = el.ExposureType,
                    Activity_Source = el.Activity_Source,
                    mediaOfRelease = el.mediaOfRelease,
                    SourceSummary = el.SourceSummary
                };
                o.GenericScenario = GetScenario(el.ESD_GS_Name);
                o.sources = GetSources(el);
                remainingValues.Add(o);
                if (!uniqueDataElements.Contains(el.ElementName)) uniqueDataElements.Add(el.ElementName);
                if (!uniqueDataSubElements.Contains(el.Type)) uniqueDataSubElements.Add(el.Type);
                remainingDataTable.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            }
            remainingValuesDataGridView.DataSource = remainingDataTable;

            ExportDataSet(genScenarios, @"..\..\output.xlsx");
        }

        void SetUpDataTables()
        {
            //foreach (string s in scenarioElements) occExpTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) procDescriptionTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) useRateTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) envReleaseTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) contolTechTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) shiftTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) operatingDaysTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) workerTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) siteTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) ppeTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) productionRateTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) parameterTable.Columns.Add(s, typeof(string));
            //foreach (string s in scenarioElements) remainingDataTable.Columns.Add(s, typeof(string));

            // dataValueDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type,el.SourceSummary });
            this.parameterTable.Columns.Add("Element Number");
            this.parameterTable.Columns.Add("Element Name");
            this.parameterTable.Columns.Add("Element Type");
            this.parameterTable.Columns.Add("Source Summary");

            // remainingValuesDataGridView1.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            this.remainingDataTable.Columns.Add("Element Number");
            this.remainingDataTable.Columns.Add("Element Name");
            this.remainingDataTable.Columns.Add("Element Type");
            this.remainingDataTable.Columns.Add("Element Type 2");
            this.remainingDataTable.Columns.Add("Activity Source");
            this.remainingDataTable.Columns.Add("Media Of Release");
            this.remainingDataTable.Columns.Add("Source Summary");


            // occupationalExposureDataGridView.Rows.Add(new string[] { o.ElementNumber, o.ScenarioName, o.ElementNumber, o.Type, o.Activity_Source, o.sourceSummary, o.mediaOfRelease, o.sourceSummary });
            this.occExpTable.Columns.Add("Element Number");
            this.occExpTable.Columns.Add("Element Name");
            this.occExpTable.Columns.Add("Element Type");
            this.occExpTable.Columns.Add("Exposure Type");
            this.occExpTable.Columns.Add("Activity Source");
            this.occExpTable.Columns.Add("Media Of Release");
            this.occExpTable.Columns.Add("Source Summary");

            // processDescriptionDataGridView.Rows.Add(new string[] { de.ElementNumber, de.ElementName, de.Type, de.Type2, de.SourceSummary
            this.procDescriptionTable.Columns.Add("Element Number");
            this.procDescriptionTable.Columns.Add("Element Name");
            this.procDescriptionTable.Columns.Add("Element Type");
            this.procDescriptionTable.Columns.Add("Element Type 2");
            this.procDescriptionTable.Columns.Add("Source Summary");

            // useRateDataGridView.Rows.Add(new string[] { ur.ElementNumber, ur.ElementName, ur.Type, ur.SourceSummary });
            this.useRateTable.Columns.Add("Element Number");
            this.useRateTable.Columns.Add("Element Name");
            this.useRateTable.Columns.Add("Element Type");
            this.useRateTable.Columns.Add("Source Summary");

            // environmentalReleaseDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.Activity_Source, el.mediaOfRelease, el.SourceSummary });
            this.envReleaseTable.Columns.Add("Element Number");
            this.envReleaseTable.Columns.Add("Element Name");
            this.envReleaseTable.Columns.Add("Element Type");
            this.envReleaseTable.Columns.Add("Element Type 2");
            this.envReleaseTable.Columns.Add("Activity_Source");
            this.envReleaseTable.Columns.Add("Media Of Release");
            this.envReleaseTable.Columns.Add("Source Summary");

            // controlTechnologyDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            this.contolTechTable.Columns.Add("Element Number");
            this.contolTechTable.Columns.Add("Element Name");
            this.contolTechTable.Columns.Add("Element Type");
            this.contolTechTable.Columns.Add("Element Type 2");
            this.contolTechTable.Columns.Add("Source Summary");

            // shiftDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            this.shiftTable.Columns.Add("Element Number");
            this.shiftTable.Columns.Add("Element Name");
            this.shiftTable.Columns.Add("Element Type");
            this.shiftTable.Columns.Add("Element Type 2");
            this.shiftTable.Columns.Add("Source Summary");

            // operatingDaysDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.Type2, el.SourceSummary });
            this.operatingDaysTable.Columns.Add("Element Number");
            this.operatingDaysTable.Columns.Add("Element Name");
            this.operatingDaysTable.Columns.Add("Element Type");
            this.operatingDaysTable.Columns.Add("Element Type 2");
            this.operatingDaysTable.Columns.Add("Source Summary");

            // workersDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.workerTable.Columns.Add("Element Number");
            this.workerTable.Columns.Add("Element Name");
            this.workerTable.Columns.Add("Element Type");
            this.workerTable.Columns.Add("Source Summary");

            // sitesDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.siteTable.Columns.Add("Element Number");
            this.siteTable.Columns.Add("Element Name");
            this.siteTable.Columns.Add("Element Type");
            this.siteTable.Columns.Add("Source Summary");

            // ppeDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.ppeTable.Columns.Add("Element Number");
            this.ppeTable.Columns.Add("Element Name");
            this.ppeTable.Columns.Add("Element Type");
            this.ppeTable.Columns.Add("Source Summary");

            // productionRateDataGridView.Rows.Add(new string[] { el.Element, el.ElementName, el.Type, el.SourceSummary });
            this.productionRateTable.Columns.Add("Element Number");
            this.productionRateTable.Columns.Add("Element Name");
            this.productionRateTable.Columns.Add("Element Type");
            this.productionRateTable.Columns.Add("Source Summary");


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
            genScenarios.Tables.Add(occExpTable);
            genScenarios.Tables.Add(envReleaseTable);
            genScenarios.Tables.Add(productionRateTable);
            genScenarios.Tables.Add(contolTechTable);
            genScenarios.Tables.Add(useRateTable);
            genScenarios.Tables.Add(siteTable);
            genScenarios.Tables.Add(operatingDaysTable);
            genScenarios.Tables.Add(workerTable);
            genScenarios.Tables.Add(shiftTable);
            genScenarios.Tables.Add(ppeTable);
            genScenarios.Tables.Add(parameterTable);
            genScenarios.Tables.Add(remainingDataTable);
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
            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadSheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(@"..\..\Revised Data Element Comparison Draft_2.19.2020_To EPA_with review notes.xlsx", false))
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
                elementsAndDataTable.Columns.Add(accessed);

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
                infoTable.Columns.Add(accessed);

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

        private void ExportDataSet(DataSet ds, string destination)
        {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
                foreach (System.Data.DataTable table in ds.Tables)
                {
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

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<String> columns = new List<string>();
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
