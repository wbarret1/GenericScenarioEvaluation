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
        DataTable elementsAndDataTable = new DataTable();
        DataTable infoTable = new DataTable();
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

        public Form1()
        {
            InitializeComponent();
            ProcessExcel();
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

            int count = 0;
            List<Source> sources = new List<Source>();
            foreach (DataRow row in elementsAndDataTable.Rows)
            {
                string name = row[scenarioElements[1]].ToString();
                for (int i = 9; i < 17; i++)
                    if (!string.IsNullOrEmpty(row[scenarioElements[i]].ToString()))
                    {
                        bool contained = false;
                        foreach (Source s in sources)
                        {
                            if (s.Value == row[scenarioElements[i]].ToString()) contained = true;
                        }
                        if (!contained)
                        {
                            sources.Add(new Source
                            {
                                ScenarioName = name,
                                Value = row[scenarioElements[i]].ToString()
                            });
                        }
                    }
            }

            this.dataGridView2.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView2.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView2.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView2.Columns.Add(scenarioElements[5], scenarioElements[5]);
            this.dataGridView2.Columns.Add(scenarioElements[6], scenarioElements[6]);
            this.dataGridView2.Columns.Add(scenarioElements[7], scenarioElements[7]);
            this.dataGridView2.Columns.Add(scenarioElements[8], scenarioElements[8]);
            this.dataGridView2.Columns.Add(scenarioElements[9], scenarioElements[9]);

            var results = from myRow in elementsAndDataTable.AsEnumerable()
                          where myRow.Field<string>("Data Element").Contains("Occupational") 
                          && !myRow.Field<string>("Data Element").Contains("Process Description")
                          select myRow;

            List<OccupationalExposure> expsoures = new List<OccupationalExposure>();
            foreach (DataRow row in results)
            {
                count++;
                expsoures.Add(new OccupationalExposure
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    ExposureType = row[scenarioElements[5]].ToString(),
                    Actiity_Source = row[scenarioElements[6]].ToString(),
                    mediaOfRelease = row[scenarioElements[7]].ToString(),
                    sourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[5]].ToString(),
                    row[scenarioElements[6]].ToString(),
                    row[scenarioElements[7]].ToString(),
                    row[scenarioElements[8]].ToString(),
                    row[scenarioElements[9]].ToString()
               };
                this.dataGridView2.Rows.Add(temp);
                row[accessed] = "True";
            }

            /*
              		[0]	"Element #"	string
                    [1]	"ESD/GS Name"	string
                    [2]	"Data Element"	string
                    [3]	"Sub-Element (Type)"	string
                    [4]	"Sub-Element (Type 2)"	string
                    [5]	"Sub-Element (Exposure Type)"	string
                    [6]	"Sub-Element (Activity/Source)"	string
                    [7]	"Sub-Element (Media of Release)"	string
                    [8]	"Data Element Source Summary"	string
                    [9]	"Data Element Source 1"	string
                    [10]	"Data Element Source 2"	string
                    [11]	"Data Element Source 3"	string
                    [12]	"Data Element Source 4"	string
                    [13]	"Data Element Source 5"	string
                    [14]	"Data Element Source 6"	string
                    [15]	"Data Element Source 7"	string
                    [16]	"Data Element Source 8"	string
                    [17]	"Reviewed"	string
                    [18]	"Reference Check"	string
                    */

            this.dataGridView3.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView3.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView3.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView3.Columns.Add(scenarioElements[4], scenarioElements[4]);
            this.dataGridView3.Columns.Add(scenarioElements[8], scenarioElements[8]);

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Process Description")
                      select myRow;

            List<ProcessDescription> processDescriptions = new List<ProcessDescription>();
            foreach (DataRow row in results)
            {
                count++;
                processDescriptions.Add(new ProcessDescription
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Description = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    Type2 = row[scenarioElements[4]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[8]].ToString()
               };
                this.dataGridView3.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            this.dataGridView4.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView4.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView4.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView4.Columns.Add(scenarioElements[8], scenarioElements[8]);

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Sub-Element (Type)").ToLower().Contains("use rate") ||
                        myRow.Field<string>("Data Element").ToLower().Contains("use rate") ||
                        myRow.Field<string>("Data Element").ToLower().Contains("daily use") ||
                        myRow.Field<string>("Data Element").ToLower().Contains("annual use")
                      select myRow;

            List<UseRate> useRates = new List<UseRate>();
            foreach (DataRow row in results)
            {
                count++;
                useRates.Add(new UseRate
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Value = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView4.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            this.dataGridView5.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView5.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView5.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView5.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView5.Columns.Add(scenarioElements[6], scenarioElements[6]);
            this.dataGridView5.Columns.Add(scenarioElements[7], scenarioElements[7]);
            this.dataGridView5.Columns.Add(scenarioElements[8], scenarioElements[8]);

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Environmental Release") ||
                      myRow.Field<string>("Data Element").Contains("TRI Releases (lb/yr)") ||
                      myRow.Field<string>("Data Element").Contains("Total Industry Estimated Process Water Discharge Flow")

                      && !myRow.Field<string>("Data Element").Contains("Process Description")
                      select myRow;

            List<EnvironmentalRelease> envRelease = new List<EnvironmentalRelease>();
            foreach (DataRow row in results)
            {
                count++;
                envRelease.Add(new EnvironmentalRelease
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    Type2 = row[scenarioElements[4]].ToString(),
                    ActiitySource = row[scenarioElements[6]].ToString(),
                    MediaOfRelease = row[scenarioElements[7]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[6]].ToString(),
                    row[scenarioElements[7]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView5.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Control Technologies") ||
                        myRow.Field<string>("Data Element").Contains("Treatment Technology")
                      select myRow;

            this.dataGridView6.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView6.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView6.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView6.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView6.Columns.Add(scenarioElements[8], scenarioElements[8]);

            List<ControlTechnology> controlTech = new List<ControlTechnology>();
            foreach (DataRow row in results)
            {
                count++;
                controlTech.Add(new ControlTechnology
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    Type2 = row[scenarioElements[4]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView6.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Shift")||
                      myRow.Field<string>(scenarioElements[3]).Contains("Shift")
                      //&& !myRow.Field<string>("Data Element").Contains("Process Description")
                      //&& !myRow.Field<string>("Data Element").Contains("Use Rate")
                      select myRow;

            this.dataGridView9.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView9.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView9.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView9.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView9.Columns.Add(scenarioElements[8], scenarioElements[8]);

            List<Shift> shifts = new List<Shift>();
            foreach (DataRow row in results)
            {
                count++;
                shifts.Add(new Shift
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    Type2 = row[scenarioElements[4]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView9.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Operating")
                      && !myRow.Field<string>("Data Element").Contains("Process Description")
                      && !myRow.Field<string>("Data Element").Contains("Use Rate")
                      select myRow;

            this.dataGridView8.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView8.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView8.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView8.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView8.Columns.Add(scenarioElements[8], scenarioElements[8]);

            List<OperatingDays> opDays = new List<OperatingDays>();
            foreach (DataRow row in results)
            {
                count++;
                opDays.Add(new OperatingDays
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    Type2 = row[scenarioElements[4]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView8.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Worker")
                      && !myRow.Field<string>("Data Element").Contains("Process Description")
                      && !myRow.Field<string>("Data Element").Contains("Use Rate")
                      select myRow;

            this.dataGridView7.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView7.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView7.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView7.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView7.Columns.Add(scenarioElements[8], scenarioElements[8]);

            List<Worker> workers = new List<Worker>();
            foreach (DataRow row in results)
            {
                count++;
                workers.Add(new Worker
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[6]].ToString(),
                    row[scenarioElements[7]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView7.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            this.dataGridView11.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView11.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView11.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView11.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView11.Columns.Add(scenarioElements[8], scenarioElements[8]);


            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("Number of Sites")
                      || myRow.Field<string>(scenarioElements[3]).Contains("Number of Sites")
                      select myRow;

            List<Site> sites = new List<Site>();
            foreach (DataRow row in results)
            {
                count++;
                sites.Add(new Site
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[6]].ToString(),
                    row[scenarioElements[7]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView11.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            this.dataGridView12.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView12.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView12.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView12.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView12.Columns.Add(scenarioElements[8], scenarioElements[8]);


            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>("Data Element").Contains("PPE")
                      || myRow.Field<string>(scenarioElements[3]).Contains("PPE")
                      select myRow;

            List<ppe> ppes = new List<ppe>();
            foreach (DataRow row in results)
            {
                count++;
                ppes.Add(new ppe
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[6]].ToString(),
                    row[scenarioElements[7]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView12.Rows.Add(temp);
                if (row[accessed].ToString() == "True")
                    System.Windows.Forms.MessageBox.Show("Found one!");
                row[accessed] = "True";
            }

            this.dataGridView10.Columns.Add(scenarioElements[1], scenarioElements[1]);
            this.dataGridView10.Columns.Add(scenarioElements[2], scenarioElements[2]);
            this.dataGridView10.Columns.Add(scenarioElements[3], scenarioElements[3]);
            this.dataGridView10.Columns.Add(scenarioElements[4], scenarioElements[5]);
            this.dataGridView10.Columns.Add(scenarioElements[8], scenarioElements[8]);

            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where !string.IsNullOrEmpty(myRow.Field<string>(scenarioElements[2]))
                      && !string.IsNullOrEmpty(myRow.Field<string>(scenarioElements[3]))
                      && myRow.Field<string>(accessed) != "True"
                      select myRow;

            List<DataValue> values = new List<DataValue>();
            List<string> uniqueDataElements = new List<string>();
            foreach (DataRow row in results)
            {
                count++;
                values.Add(new DataValue
                {
                    ScenarioName = row[scenarioElements[1]].ToString(),
                    Element = row[scenarioElements[2]].ToString(),
                    Type = row[scenarioElements[3]].ToString(),
                    Type2 = row[scenarioElements[4]].ToString(),
                    SourceSummary = row[scenarioElements[8]].ToString()
                });
                if (!uniqueDataElements.Contains(row[scenarioElements[2]].ToString())) uniqueDataElements.Add(row[scenarioElements[2]].ToString());
                string[] temp = new string[] {
                    row[scenarioElements[1]].ToString(),
                    row[scenarioElements[2]].ToString(),
                    row[scenarioElements[3]].ToString(),
                    row[scenarioElements[4]].ToString(),
                    row[scenarioElements[8]].ToString()
                };
                this.dataGridView10.Rows.Add(temp);
                row[accessed] = "True";
            }


            results = from myRow in elementsAndDataTable.AsEnumerable()
                      where myRow.Field<string>(accessed) != "True"
                      select myRow;

            int remaining = results.Count();
            foreach (string s in scenarioElements)
            {
                dataGridView1.Columns.Add(s, s);
            }
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            foreach (DataRow r in results)
            {
                List<string> data = new List<string>();
                for (int i = 0; i < scenarioElements.Length; i++) data.Add(r[scenarioElements[i]].ToString());
                this.dataGridView1.Rows.Add(data.ToArray<string>());
            }

            // format data grid views
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView3.Columns.Count; i++)
            {
                dataGridView3.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView4.Columns.Count; i++)
            {
                dataGridView4.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView5.Columns.Count; i++)
            {
                dataGridView5.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView6.Columns.Count; i++)
            {
                dataGridView6.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView7.Columns.Count; i++)
            {
                dataGridView7.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView8.Columns.Count; i++)
            {
                dataGridView8.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView9.Columns.Count; i++)
            {
                dataGridView9.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dataGridView10.Columns.Count; i++)
            {
                dataGridView10.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }


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
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
    }
}
