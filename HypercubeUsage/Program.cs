using System;
using System.Linq;
using HypercubeUsage.Properties;
using Qlik.Engine;
using Qlik.Sense.Client;
using Qlik.Sense.Client.Visualizations;

namespace HypercubeUsage
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            using (var hub = ConnectToDesktop().Hub())
            {
                var theApp = PrepareApp(hub);
                HyperCubeExamples(theApp);
                Paging(theApp);
                MultipleHyperCubesExamples(theApp);
            }
        }

        private static IApp PrepareApp(IHub hub)
        {
            var theApp = hub.CreateSessionApp().Return;
            theApp.SetScript(Resources.LoadScript);
            theApp.CreateMeasure("Sales",
                new MeasureProperties
                {
                    Measure = new NxLibraryMeasureDef {Def = "Sum([Sales Amount])"},
                    MetaDef = new MetaAttributesDef {Title = "Sales"}
                }
                );
            theApp.DoReload();
            return theApp;
        }

        private static ILocation ConnectToDesktop()
        {
            var location = Location.FromUri(new Uri("ws://127.0.0.1:4848"));
            location.AsDirectConnectionToPersonalEdition();
            location.IsVersionCheckActive = false;
            return location;
        }

        private static IGenericObject CreateCustomObject(IApp theApp)
        {
            // Creates a session object with a custom type.
            var customObjectProperties = new GenericObjectProperties { Info = new NxInfo { Type = "mycustomtype" } };
            return theApp.CreateGenericSessionObject(customObjectProperties);
        }

        private static void SetHyperCube(IGenericObject theObject, HyperCubeDef hyperCube)
        {
            using (theObject.SuspendedLayout)
            {
                // A property with the name "qHyperCubeDef" will be interpreted by the engine as being of type HyperCubeDef
                // and evaluated accordingly.
                theObject.Properties.Set("qHyperCubeDef", hyperCube);
            }
        }

        private static void HyperCubeExamples(IApp theApp)
        {
            // Add an empty hypercube to the newly created object.
            var theObject = CreateCustomObject(theApp);
            SetHyperCube(theObject, new HyperCubeDef());

            BasicDimensionAndMeasureUsage(theApp, theObject);
            ConfigureFormating(theObject);
            ConfigureSorting(theObject);
            BasicSelection(theApp, theObject);
            GroupedDimensions(theApp, theObject);
        }

        private static void BasicDimensionAndMeasureUsage(IApp theApp, IGenericObject theObject)
        {
            var theHyperCube = theObject.Properties.Get<HyperCubeDef>("qHyperCubeDef");

            // List Sales per year (1 Dimension, 1 Measure)
            AddInlineDimension(theHyperCube, "Year");
            AddInlineMeasure(theHyperCube, "Sum([Sales Amount])");
            SetHyperCube(theObject, theHyperCube);
            // Columns of data appear in the order they appear in the dimension and measure definition lists
            // for the hypercube. Dimensions first, then measures. Column 0 will therefore contain the years,
            // and column 1 will contain the sum of sales.
            PrintData("Sales per year", theObject,
                row => String.Format("Year: {0}, Sales: {1}", row[0].Num, row[1].Num));

            // List sales per year and month
            AddInlineDimension(theHyperCube, "Month");
            SetHyperCube(theObject, theHyperCube);
            // Adding a dimension increases the number of columns in the returned data. The new dimension
            // will appear on position 1 (after the year) as AddInlineDimension appends dimensions to the
            // end of the dimension list.
            PrintData("Sales per year and month", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Num, row[2].Num));

            // Use predefined library version of the measure.
            // Get the ID of the measure with the title "Sales" by looking up the measures in the measure list.
            theHyperCube.Measures = Enumerable.Empty<NxMeasure>();
            var theMeasureList = theApp.GetMeasureList();
            var measureId = theMeasureList.Items.First(item => item.Data.Title == "Sales").Info.Id;
            AddLibraryMeasure(theHyperCube, measureId);
            SetHyperCube(theObject, theHyperCube);
            PrintData("Sales per year and month using library measure for sales", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Num, row[2].Num));
        }

        // An inline dimension is a dimension that is directly defined in the HyperCubeDef structure.
        // The alternative is to use dimensions predefined in the library.
        private static void AddInlineDimension(HyperCubeDef theHyperCube, string field)
        {
            var inlineDimension = new NxInlineDimensionDef { FieldDefs = new[] { field } };
            // A dimension directly defined in the hypercube definition. The alternative would be to
            // refer to a predefined library dimension in which case the property LibraryId would be used instead.
            var dimension = new NxDimension { Def = inlineDimension };
            theHyperCube.Dimensions = theHyperCube.Dimensions.Concat(new[] { dimension });
        }

        private static void AddInlineMeasure(HyperCubeDef theHyperCube, string expression)
        {
            var inlineMeasure = new NxInlineMeasureDef {Def = expression};
            // Analogous to AddInlineDimension. This NxMeasure represents a measure directly defined in the hyper
            // cube definition. The alternative would be to refer to a predefined library dimension in which case
            // the property LibraryId would be used instead.
            var measure = new NxMeasure {Def = inlineMeasure};
            theHyperCube.Measures = theHyperCube.Measures.Concat(new[] { measure });
        }

        // Measures can be predefined in the library and thereby reused in in multiple hypercubes. The alternative
        // is to define the measure directly in the hypercube.
        private static void AddLibraryMeasure(HyperCubeDef theHyperCube, string measureId)
        {
            // A library measure is referred to by the id used when creating it.
            var measure = new NxMeasure { LibraryId = measureId };
            theHyperCube.Measures = theHyperCube.Measures.Concat(new[] { measure });
        }

        private static void ConfigureFormating(IGenericObject theObject)
        {
            // Use string literal for month by using the Text representation instead of the numeric representation
            // for the month column.
            PrintData("String literal for months", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Num));

            var hyperCube = theObject.Properties.Get<HyperCubeDef>("qHyperCubeDef");
            // Format sales as USD
            var measure = hyperCube.Measures.Single();
            measure.Def.NumFormat = new FieldAttributes { Type = FieldAttrType.MONEY, nDec = 2, Dec = ".", UseThou = 1, Thou = "," };
            SetHyperCube(theObject, hyperCube);
            PrintData("Use USD as currency", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Text));
        }

        private static void ConfigureSorting(IGenericObject theObject)
        {
            var hyperCube = theObject.Properties.Get<HyperCubeDef>("qHyperCubeDef");

            // List sales per year and month, Sort on Year and Month, both ascending
            {
                var dimensions = hyperCube.Dimensions;
                foreach (var dimension in dimensions)
                {
                    // Sort both dimensions ascending.
                    dimension.Def.SortCriterias = new[] { new SortCriteria { SortByNumeric = SortDirection.Ascending } };
                }
                SetHyperCube(theObject, hyperCube);
            }
            PrintData("Sales per year and month, sorted", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Text));

            // Sort by month first, then year, then sales.
            hyperCube.InterColumnSortOrder = new[] { 1, 0, 2 };
            SetHyperCube(theObject, hyperCube);

            PrintData("Sales per year and month, sorted by month then year", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Num, row[2].Text));

            // Sort by month first, then year, then sales.
            hyperCube.InterColumnSortOrder = new[] { 1, 0, 2 };
            SetHyperCube(theObject, hyperCube);

            PrintData("Sales per year and month, sorted by month then year", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Num, row[2].Text));

            // Sort by sales (descending), then year, then month.
            hyperCube.InterColumnSortOrder = new[] { 2, 0, 1 };
            hyperCube.Measures.First().SortBy = new SortCriteria {SortByNumeric = SortDirection.Descending};
            SetHyperCube(theObject, hyperCube);

            PrintData("Sales per year and month, sorted by sales (descending)", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Num, row[2].Text));

            // Revert to sort by year first.
            hyperCube.InterColumnSortOrder = new[] { 0, 1, 2 };
            SetHyperCube(theObject, hyperCube);
        }

        private static void BasicSelection(IApp theApp, IGenericObject theObject)
        {
            // Print data for year 2016 only.
            theApp.GetField("Year").Select("2016");
            PrintData("Sales for year 2016 only", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Text));

            var salesRepField = theApp.GetField("Sales Rep Name");
            // Print data for year 2016 only, and for sales rep "Amalia Craig" only.
            salesRepField.Select("Amalia Craig");
            PrintData("Sales for year 2016 and sales rep Amalia Craig", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Text));

            // Switch to sales rep "Amanda Honda"
            salesRepField.Clear();
            salesRepField.Select("Amanda Honda");
            // Print data for year 2016 only, and for sales rep "Amanda Honda" only.
            PrintData("Sales for year 2016 and sales rep Amanda Honda", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Text));

            // Clear selections and print data for all years and sales reps.
            theApp.ClearAll();
            PrintData("Sales for all years", theObject,
                row => String.Format("Year: {0}, Month: {1}, Sales: {2}", row[0].Num, row[1].Text, row[2].Text));

        }

        private static void GroupedDimensions(IApp theApp, IGenericObject theObject)
        {
            var hyperCube = theObject.Properties.Get<HyperCubeDef>("qHyperCubeDef");

            // Make YearMonth, Year and Month stacked dimensions
            hyperCube.Dimensions = Enumerable.Empty<NxDimension>();
            AddInlineDimension(hyperCube, "Year");
            var dimension = hyperCube.Dimensions.ToArray()[0];
            dimension.Def.FieldDefs = dimension.Def.FieldDefs.Concat(new[] { "Month" });
            // The Grouping property indicates that the Dimension is stacked.
            dimension.Def.Grouping = NxGrpType.GRP_NX_HIEARCHY;
            dimension.Def.SortCriterias = new[]
            {
                new SortCriteria {SortByNumeric = SortDirection.Ascending},
                new SortCriteria {SortByNumeric = SortDirection.Ascending}
            };
            SetHyperCube(theObject, hyperCube);

            // Prints sales per Year
            PrintGroupedData("Stacked, Sales per Year", theObject);
            // When selecting a single field of the top dimension of the stack, the hypercube
            // will use the next dimension of the stack instead.
            theApp.GetField("Year").Select("2016");
            // Prints sales per month for the selected year.
            PrintGroupedData("Stacked, Sales per Month for selected Year 2016", theObject);

            // Clear year selection.
            theApp.ClearAll();
            // Set group mode to cyclic and add dimension YearMonth
            dimension.Def.Grouping = NxGrpType.GRP_NX_COLLECTION;
            dimension.Def.FieldDefs = dimension.Def.FieldDefs.Concat(new[] { "YearMonth" });
            SetHyperCube(theObject, hyperCube);
            // Print data for Year
            PrintGroupedData("Cyclic, Sales per Year (group 0)", theObject);
            // Print data for Month
            dimension.Def.ActiveField = 1;
            SetHyperCube(theObject, hyperCube);
            PrintGroupedData("Cyclic, Sales per Month (group 1)", theObject);
            // Print data for YearMonth
            dimension.Def.ActiveField = 2;
            SetHyperCube(theObject, hyperCube);
            PrintGroupedData("Cyclic, Sales per YearMonth (group 2)", theObject);
        }

        private static void Paging(IApp theApp)
        {
            var theObject = CreateCustomObject(theApp);

            // Create a hypercube with sales per YearMonth.
            var salesPerYearMonthHc = new HyperCubeDef();
            AddInlineDimension(salesPerYearMonthHc, "YearMonth");
            AddInlineMeasure(salesPerYearMonthHc, "Sum([Sales Amount])");
            SetHyperCube(theObject, salesPerYearMonthHc);

            // Get the hypercube pager
            var pager = theObject.GetHyperCubePager("/qHyperCubeDef");

            // Get data for a page with width 2 and height 5:
            var firstPage = new[] { new NxPage { Top = 0, Left = 0, Width = 2, Height = 5 } };
            var firstData = pager.GetData(firstPage).Single();
            PrintPage("First page:", firstData,  row => String.Format("YearMonth: {0}, Sales: {1}", row[0].Text, row[1].Text));

            // Print all other pages:
            var pageNr = 1;
            while (!pager.OutsideEdge.Single())
            {
                var nextData = pager.GetNextPage().Single();
                PrintPage("Page nr: " + pageNr, nextData, row => String.Format("YearMonth: {0}, Sales: {1}", row[0].Text, row[1].Text));
                pageNr++;
            }

            // Iterate across all pages:
            pageNr = 0;
            firstPage = new[] { new NxPage { Top = 0, Left = 0, Width = 2, Height = 5 } };
            foreach(var page in pager.IteratePages(firstPage, Pager.Next))
            {
                PrintPage("Page nr: " + pageNr, page.First(), row => String.Format("YearMonth: {0}, Sales: {1}", row[0].Text, row[1].Text));
                pageNr++;                
            }

            // Accessing pager for client object hypercubes:
            var theSheet = GetOrCreateSheet(theApp, "mySheet");
            var barChart = theSheet.CreateBarchart();
            // Set hypercube to sales per year and month
            using (barChart.SuspendedLayout)
            {
                barChart.Properties.HyperCubeDef = salesPerYearMonthHc.CloneAs<VisualizationHyperCubeDef>();
            }
            // Print first page of bar chart data:
            var firstBarchartData = barChart.HyperCubePager.GetData();
            PrintPage("First page of bar chart:", firstBarchartData.Single(),
                row => String.Format("YearMonth: {0}", row[0].Text)
                );
        }

        private static ISheet GetOrCreateSheet(IApp theApp, string sheetId)
        {
            return theApp.GetSheet("sheetId") ?? theApp.CreateSheet(sheetId);
        }

        // Print all data for the hypercube. The argument function is used to print individual rows of data.
        private static void PrintData(string header, IGenericObject theObject, Func<NxCellRows, string> format, string hyperCubePath = "/qHyperCubeDef")
        {
            // Get the data for the hypercube. 
            var theData = theObject.GetHyperCubeData(hyperCubePath, new[] { LargePage }).First();
            // Print the page using the format function.
            PrintPage(header, theData, format);
        }

        private static void PrintPage(string header, NxDataPage theData, Func<NxCellRows, string> format)
        {
            Console.WriteLine("*** " + header);
            foreach (var row in theData.Matrix)
            {
                // The column count is equal to the sum of the number of dimensions and the number of measures.
                Console.WriteLine(format(row));
            }
        }

        // Represents a page large enough to fit all data used in the example. See section on paging for examples
        // of how to work with pages and page sizes.
        public static NxPage LargePage
        {
            get { return new NxPage {Top = 0, Left = 0, Width = 10, Height = 100}; }
        }

        private static void PrintGroupedData(string header, IGenericObject theObject)
        {
            var hyperCubeLayout = theObject.GetLayout().Get<HyperCube>("qHyperCube");
            // The hypercube has only a single dimension. Which dimension is active depends on the state
            // of the hypercube.
            var dimensionInfo = hyperCubeLayout.DimensionInfo.Single();
            var template = dimensionInfo.FallbackTitle + ": {0}, Sales: {1}";
            var isYear = dimensionInfo.GroupPos == 0;
            // If the current group position is 0, then the dimension represents a year, if it is 1, then it
            // represents a month. Use Num property of cell for Year and Text for month.
            PrintData(header, theObject,
                row => String.Format(template, isYear ? (object) row[0].Num : row[0].Text, row[1].Text));
        }

        private static void MultipleHyperCubesExamples(IApp theApp)
        {
            var theObject = CreateCustomObject(theApp);
            // Create two hypercubes with the same measure, but different dimensions.
            var salesPerMonthHc = new HyperCubeDef();
            AddInlineDimension(salesPerMonthHc, "Month");
            AddInlineMeasure(salesPerMonthHc, "Sum([Sales Amount])");
            var salesPerYearHc = new HyperCubeDef();
            AddInlineDimension(salesPerYearHc, "Year");
            AddInlineMeasure(salesPerYearHc, "Sum([Sales Amount])");

            // Add the hypercubes to containers. The containers are used to separate the two
            // qHyperCubeDef properties in the property tree of the object.
            var hcContainer0 = new DynamicStructure();
            hcContainer0.Set("qHyperCubeDef", salesPerMonthHc);
            var hcContainer1 = new DynamicStructure();
            hcContainer1.Set("qHyperCubeDef", salesPerYearHc);

            // Add containers to object.
            using (theObject.SuspendedLayout)
            {
                theObject.Properties.Set("container0", hcContainer0);
                theObject.Properties.Set("container1", hcContainer1);
            }

            // Print data for first hypercube
            PrintData("Sales per Month", theObject,
                row => String.Format("Month: {0}, Sales: {1}", row[0].Text, row[1].Text),
                // Path to hypercube 0. (Name of container property followed by hypercube property.)
                "/container0/qHyperCubeDef");

            // Print data for second hypercube
            PrintData("Sales per Year", theObject,
                row => String.Format("Year: {0}, Sales: {1}", row[0].Num, row[1].Text),
                // Path to hypercube 1. (Name of container property followed by hypercube property.)
                "/container1/qHyperCubeDef");

            // Get pager for cube in container 1:
            var thePager = theObject.GetAllHyperCubePagers().First(pager => pager.Path.Contains("container0"));
            // Get Last five rows for that hypercube:
            thePager.CurrentPages = new []{new NxPage{Width = 2, Height = 5}};
            var theLastFiveRows = thePager.GetLastPage().First();
            PrintPage("The last five rows of hypercube in container0:", theLastFiveRows,
                row => String.Format("Month: {0}, Sales: {1}", row[0].Text, row[1].Text)
                );
        }
    }
}
