using System.Configuration;

using GNAgeneraltools;

using GNAspreadsheettools;

using OfficeOpenXml;

//using Twilio;
//using Twilio.Rest.Api.V2010.Account;
//using Twilio.Types;


namespace sortRails
{
    class Program
    {
        static void Main(string[] args)
        {
            //===============[Suppress warnings]======================================
#pragma warning disable CS0162
#pragma warning disable CS0164
#pragma warning disable CS0168
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416

            // Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;

            // Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // instantiate the classes
            spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI();
            gnaTools gnaT = new gnaTools();


            gnaT.WelcomeMessage("sortRails program");

            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];

            string strActiveWorksheet = ConfigurationManager.AppSettings["ActiveWorksheet"];

            string strNameColumn = ConfigurationManager.AppSettings["NameColumn"];
            string strFirstDataRow = ConfigurationManager.AppSettings["FirstDataRow"];
            string strFreezeScreen = ConfigurationManager.AppSettings["FreezeScreen"];
            string strLeftRailTag = ConfigurationManager.AppSettings["LeftRailTag"];
            string strRightRailTag = ConfigurationManager.AppSettings["RightRailTag"];
            string strWorkbookFullPath = strExcelPath + strExcelFile;

            String[] strTag = new String[20];

            strTag[0] = ConfigurationManager.AppSettings["LeftRailTag1"];
            strTag[1] = ConfigurationManager.AppSettings["RightRailTag1"];
            strTag[2] = ConfigurationManager.AppSettings["LeftRailTag2"];
            strTag[3] = ConfigurationManager.AppSettings["RightRailTag2"];
            strTag[4] = ConfigurationManager.AppSettings["LeftRailTag3"];
            strTag[5] = ConfigurationManager.AppSettings["RightRailTag3"];
            strTag[6] = ConfigurationManager.AppSettings["LeftRailTag4"];
            strTag[7] = ConfigurationManager.AppSettings["RightRailTag4"];
            strTag[8] = ConfigurationManager.AppSettings["LeftRailTag5"];
            strTag[9] = ConfigurationManager.AppSettings["RightRailTag5"];
            strTag[10] = ConfigurationManager.AppSettings["LeftRailTag6"];
            strTag[11] = ConfigurationManager.AppSettings["RightRailTag6"];

            //==== Environment check

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("     Workbook: " + strWorkbookFullPath);
            gnaSpreadsheetAPI.checkWorksheetExists(strWorkbookFullPath, strActiveWorksheet);

            int iRow = 0;
            int i = 0;
            int j = 0;
            int iCol = 0;
            int iPrismCounter = 0;
            string strName = "";
            string strFirstPass = "Yes";

            double dblE, dblN, dblH;
            string ATS;

            var prism = new List<Prism>();
            var sortedLeftRailprisms = new List<Prism>();
            var sortedRightRailprisms = new List<Prism>();
            int iLeftRailCounter = -1;
            int iRightRailCounter = -1;
            var strRailFlag = "LeftRail";

            // read all prisms Left Rail Tag
            FileInfo newFile = new FileInfo(strWorkbookFullPath);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets[strActiveWorksheet];

                Console.WriteLine("2. Sort rails");
                iRow = Convert.ToInt16(strFirstDataRow);

                iPrismCounter = -1;
                strName = "blank";
                iCol = Convert.ToInt32(strNameColumn);

                j = 0;
                do
                {
                    Console.WriteLine("   Tag: "+ strTag[j]);
                    do
                    {
                        strName = Convert.ToString(worksheet.Cells[iRow, iCol].Value);

                        if (strName.Contains(strTag[j]))
                        {
                            iPrismCounter++;
                            dblE = Convert.ToDouble(worksheet.Cells[iRow, iCol + 1].Value);
                            dblN = Convert.ToDouble(worksheet.Cells[iRow, iCol + 2].Value);
                            dblH = Convert.ToDouble(worksheet.Cells[iRow, iCol + 3].Value);
                            ATS = Convert.ToString(worksheet.Cells[iRow, iCol + 4].Value);
                            prism.Add(new Prism() { Name = strName, E = dblE, N = dblN, H = dblH, ATS = ATS });
                        }

                        iRow++;
                        strName = Convert.ToString(worksheet.Cells[iRow, iCol].Value);

                    } while (strName != "");

                    prism.Add(new Prism() { Name = "TheEnd", E = 0, N = 0, H = 0, ATS = "TheEnd" });

                    prism.Sort(delegate (Prism x, Prism y)
                    {
                        return x.Name.CompareTo(y.Name);
                    });

                    if (strRailFlag == "RightRail")
                    {
                        strRailFlag = "LeftRail";
                        i = 0;
                        do
                        {
                            strName = prism[i].Name;
                            dblE = prism[i].E;
                            dblN = prism[i].N;
                            dblH = prism[i].H;
                            ATS = prism[i].ATS;
                            sortedRightRailprisms.Add(new Prism() { Name = strName, E = dblE, N = dblN, H = dblH, ATS = ATS });
                            i++;
                        } while (prism[i].Name.Trim() != "TheEnd");

                        sortedRightRailprisms.Add(new Prism() { Name = "EoR=============================", E = 0, N = 0, H = 0, ATS = "EoR" });
                        prism = new List<Prism>();
                        iRow = Convert.ToInt16(strFirstDataRow);
                        iPrismCounter = -1;
                        j++;
                    }
                    else
                    {
                        strRailFlag = "RightRail";
                        i = 0;
                        do
                        {
                            strName = prism[i].Name;
                            dblE = prism[i].E;
                            dblN = prism[i].N;
                            dblH = prism[i].H;
                            ATS = prism[i].ATS;
                            sortedLeftRailprisms.Add(new Prism() { Name = strName, E = dblE, N = dblN, H = dblH, ATS = ATS });
                            i++;
                        } while (prism[i].Name.Trim() != "TheEnd");

                        sortedLeftRailprisms.Add(new Prism() { Name = "EoR=============================", E = 0, N = 0, H = 0, ATS = "EoR" });
                        prism = new List<Prism>();
                        iRow = Convert.ToInt16(strFirstDataRow);
                        iPrismCounter = -1;
                        j++;
                    }

                } while (strTag[j] != "None");

            } 

            sortedLeftRailprisms.Add(new Prism() { Name = "TheEnd", E = 0, N = 0, H = 0, ATS = "TheEnd" });
            sortedRightRailprisms.Add(new Prism() { Name = "TheEnd", E = 0, N = 0, H = 0, ATS = "TheEnd" });

            Console.WriteLine("3. Write data");
            i = 0;
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                iRow = Convert.ToInt16(strFirstDataRow);
                iCol = Convert.ToInt32(strNameColumn) + 7;
                i = 0;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[strActiveWorksheet];
                do
                {
                    worksheet.Cells[iRow, iCol].Value = sortedLeftRailprisms[i].Name;
                    worksheet.Cells[iRow, iCol+1].Value = sortedLeftRailprisms[i].E;
                    worksheet.Cells[iRow, iCol+2].Value = sortedLeftRailprisms[i].N;
                    worksheet.Cells[iRow, iCol+3].Value = sortedLeftRailprisms[i].H;
                    worksheet.Cells[iRow, iCol+4].Value = sortedLeftRailprisms[i].ATS;

                    int iNextRow = iRow + 1;

                    // string strFormula = "=((j" + Convert.ToString(iRow) + "-j" + Convert.ToString(iNextRow) + ")^2+(k" + Convert.ToString(iRow) + "-k" + Convert.ToString(iNextRow) + + ")^2)^0.5"


                    string strFormula = "=ROUND(((J" + Convert.ToString(iRow) + "-J" + Convert.ToString(iNextRow) + ")^2+(K" + Convert.ToString(iRow) + "-K" + Convert.ToString(iNextRow) + ")^2)^0.5,2)";
                    worksheet.Cells[iRow, iCol + 5].Formula = strFormula;

                    strFormula = "=ROUND(((J"+ iRow +"-Q" + iRow +")^2+(K" +iRow+ "-R" +iRow+ ")^2)^0.5,2)";
                    worksheet.Cells[iRow, iCol + 6].Formula = strFormula;

                    iRow++;
                    i++;
                } while (sortedLeftRailprisms[i].Name.Trim() != "TheEnd");

                iRow = Convert.ToInt16(strFirstDataRow);
                iCol = Convert.ToInt32(strNameColumn) + 14;
                i = 0;
                do
                {
                    worksheet.Cells[iRow, iCol].Value = sortedRightRailprisms[i].Name;
                    worksheet.Cells[iRow, iCol + 1].Value = sortedRightRailprisms[i].E;
                    worksheet.Cells[iRow, iCol + 2].Value = sortedRightRailprisms[i].N;
                    worksheet.Cells[iRow, iCol + 3].Value = sortedRightRailprisms[i].H;
                    worksheet.Cells[iRow, iCol + 4].Value = sortedRightRailprisms[i].ATS;

                    int iNextRow = iRow + 1;

                    string strFormula = "=ROUND(((Q" + Convert.ToString(iRow) + "-Q" + Convert.ToString(iNextRow) + ")^2+(R" + Convert.ToString(iRow) + "-R" + Convert.ToString(iNextRow) + ")^2)^0.5,2)";
                    worksheet.Cells[iRow, iCol + 5].Formula = strFormula;


                    strFormula = "=ROUND((L" + Convert.ToString(iNextRow) + "-L" + Convert.ToString(iRow) + "),3)";
                    worksheet.Cells[iRow, iCol + 6].Formula = strFormula;
                    strFormula = "=ROUND((S" + Convert.ToString(iNextRow) + "-S" + Convert.ToString(iRow) + "),3)";
                    worksheet.Cells[iRow, iCol + 7].Formula = strFormula;

                    iRow++;
                    i++;
                } while (sortedRightRailprisms[i].Name.Trim() != "TheEnd");

                try
    
                {
                    worksheet.Calculate();
                    package.Save();
                    package.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("");
                    Console.WriteLine("Error:");
                    Console.WriteLine("sortRails: " + strWorkbookFullPath);
                    Console.WriteLine("\nClose the workbook and re-run.");
                    Console.WriteLine("");
                    Console.WriteLine(ex);
                    Console.WriteLine("");
                    Console.WriteLine("Press any key to exit..");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

            }

            ThatsAllFolks:


            Console.WriteLine("\nRails sorted & stored ...");

            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine("\nfreezeScreen set to Yes");
                Console.WriteLine("press key to exit..");
                Console.ReadKey();
            }

            Environment.Exit(0);

        }
    }
}

