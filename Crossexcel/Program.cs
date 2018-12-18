using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.Data;
using Microsoft.Win32;
using System.Data.SqlClient;

namespace Crossexcel
{
    class Program
    {
        public static byte intEMPRESAID = 0;
        public static byte intSUCURSALID = 0;

        public const string appName = "GrupoGuadiana";
        public const string section = "Config";

        static void Main(string[] args)
        {
            intEMPRESAID = Convert.ToByte(GetSetting(appName, section, "EmpresaID", String.Empty));
            intSUCURSALID = Convert.ToByte(GetSetting(appName, section, "SucursalID", String.Empty));

            DateTime date = DateTime.Now;
            string datewithformat = date.ToString();
            string dateday = date.ToString("dd MMMM yyyy HH mm ");

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;
                //IStyle headerStyle = wo.Styles.Add("HeaderStyle");
                IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.EnableSheetCalculations();
                DataTable tabla = GetExistenciaDataTable();
                int osos = tabla.Rows.Count;

                worksheet.ImportDataTable(tabla, true, 2,1);
                worksheet.AutoFilters.FilterRange = worksheet.Range["A2:H2"];
                string namesuc = GetName(intSUCURSALID);
                worksheet.Range["A1"].Text = "Llantas y Rines del Guadiana S.A. de C.V. - Existencias LRG Al " + dateday + namesuc;
                // worksheet.Range["A1"].Text = "Llantas y Rines del Guadiana S.A. de C.V. - Existencias LRG Al "+dateday+"- B4 Francisco Villa";
                worksheet.Rows[1].FreezePanes();
                worksheet.Rows[2].FreezePanes();

                IStyle headerStyle = workbook.Styles.Add("HeaderStyle");
                headerStyle.BeginUpdate();

                workbook.SetPaletteColor(8, System.Drawing.Color.FromArgb(46, 204, 113));

                headerStyle.Color = System.Drawing.Color.FromArgb(46, 204, 113);

                headerStyle.Font.Bold = true;

                headerStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;

                headerStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;

                headerStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;

                headerStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;

                headerStyle.EndUpdate();

                worksheet.Rows[1].CellStyle = headerStyle;

                IStyle pStyle = workbook.Styles.Add("pStyle");
                pStyle.BeginUpdate();

                workbook.SetPaletteColor(9, System.Drawing.Color.FromArgb(89, 171, 227));

                pStyle.Color = System.Drawing.Color.FromArgb(89, 171, 227);

                pStyle.Font.Bold = true;

                pStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;

                pStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;

                pStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;

                pStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;

                pStyle.EndUpdate();

                worksheet.Rows[0].CellStyle = pStyle;
                worksheet.SetColumnWidth(1, 13);
                worksheet.SetColumnWidth(2, 50);
                worksheet.SetColumnWidth(3,17);
                worksheet.SetColumnWidth(4, 28);

                // Create Table with data in the given range
                int soviet = osos;
                int rojos = soviet + 3;
                int rus = soviet + 4;
                string rusia = rus.ToString();
                string cossacks = rojos.ToString();
                string gulag = "A2:H" + cossacks + "";
                //IListObject table = worksheet.ListObjects.Create("Table1", worksheet[gulag]);
                string registros = soviet.ToString();
                //IRange range = worksheet.Range[gulag];
                //table.ShowTotals = true;
                //table.Columns[0].TotalsRowLabel = "Total";

                //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                string chorchill = "H2,H" + cossacks + "";
                string russevel = "H" + rusia + "";
                string totalr = "A" + rusia + "";
                worksheet.Range[totalr].Text = registros+" Registros";
                worksheet.Range[totalr].CellStyle = pStyle;
                string nrusia = "=SUM(H2:H" + cossacks + ")";
                worksheet.Range[russevel].Formula = nrusia;
                worksheet.Range[russevel].CellStyle = pStyle;
                //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                //table.Columns[5].TotalsCalculation = ExcelTotalsCalculation.Sum;
                //hacer el subtotal pero conformula ** el otro marca error con total calculation 
                //range.SubTotal(0, ConsolidationFunction.Sum, new int[] {1,rojos});
                string namer = dateday;
                string fileName = "LRG-Existencias al " + namer + namesuc+".xlsx";
                // string fileName = "LRG-Existencias al " + namer + "B4 Francisco Villa.xlsx";
                workbook.SaveAs(fileName);
                workbook.Close();
                excelEngine.Dispose();

            }


        }
        static DataTable GetExistenciaDataTable()
        {
            string conect = "SERVER = gggctserver.database.windows.net; DATABASE = rdbms_GGGC_Public_TESTING; USER ID = abril; PASSWORD = gggc.2017";

            SqlConnection sqlconn = new SqlConnection(conect);
            sqlconn.Open();
            string oso = GetVista(intSUCURSALID);
            SqlDataAdapter adapter = new SqlDataAdapter(oso, sqlconn);

            //SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM dbo.vtaExitencia44  ", sqlconn);
            DataSet dsPubs = new DataSet("Pubs");
            adapter.Fill(dsPubs, "Existencias_Sucursales");
            DataTable dtbl = new DataTable();

            dtbl = dsPubs.Tables["Existencias_Sucursales"];
            sqlconn.Close();
            //tratando de agrgar una columna de un data set ya definido 

            return dtbl;
        }

        static string GetVista(int empresa)
        {
            string conexion;
            switch (empresa)
            {
                case 1:
                    conexion = "SELECT * FROM dbo.vtaExitencia1  "; ;
                    return conexion;
                    break;

                case 2:
                    conexion = "SELECT * FROM dbo.vtaExitencia2  "; ;
                    return conexion;
                    break;
                case 3:
                    conexion = "SELECT * FROM dbo.vtaExitencia3  "; ;
                    return conexion;
                    break;

                case 4:
                    conexion = "SELECT * FROM dbo.vtaExitencia44  ";
                    return conexion;
                    break;
                case 5:
                    conexion = "SELECT * FROM dbo.vtaExitencia5  ";
                    return conexion;
                    break;

                case 7:
                    conexion = "SELECT * FROM dbo.vtaExitencia7  "; ;
                    return conexion;
                    break;
                case 8:
                    conexion = "SELECT * FROM dbo.vtaExitencia8  "; ;
                    return conexion;
                    break;

                case 9:
                    conexion = "SELECT * FROM dbo.vtaExitencia9  ";
                    return conexion;
                    break;
                case 10:
                    conexion = "SELECT * FROM dbo.vtaExitencia10  ";
                    return conexion;
                    break;

                case 11:
                    conexion = "SELECT * FROM dbo.vtaExitencia11  ";
                    return conexion;
                    break;
                case 12:
                    conexion = "SELECT * FROM dbo.vtaExitencia12  ";
                    return conexion;
                    break;

                case 13:
                    conexion = "SELECT * FROM dbo.vtaExitencia13  ";
                    return conexion;
                    break;
                case 15:
                    conexion = "SELECT * FROM dbo.vtaExitencia15  ";
                    return conexion;
                    break;

                case 16:
                    conexion = "SELECT * FROM dbo.vtaExitencia16  ";
                    return conexion;
                    break;
                case 17:
                    conexion = "SELECT * FROM dbo.vtaExitencia17  ";
                    return conexion;
                    break;

                case 18:
                    conexion = "SELECT * FROM dbo.vtaExitencia18  ";
                    return conexion;
                    break;
                default:
                    conexion = "SERVER = 192.168.14.1; DATABASE = Punto_De_Venta; USER ID = sa; PASSWORD = dgo2007";
                    return conexion;
                    break;

            }
        }
        static string GetName(int empresa)
        {
            string conexion;
            switch (empresa)
            {
                case 1:
                    conexion = "- B1 Felipe Pescador";
                    return conexion;
                    break;

                case 2:
                    conexion = "- B2 Francisco Villa 2" ;
                    return conexion;
                    break;
                case 3:
                    conexion = "- B3 Vizcaya";
                    return conexion;
                    break;

                case 4:
                    conexion = "- B4 Francisco Villa";
                    return conexion;
                    break;
                case 5:
                    conexion = "- B5 Blvd. Durango";
                    return conexion;
                    break;

                case 7:
                    conexion = "- B7 Mayoreo";
                    return conexion;
                    break;
                case 8:
                    conexion = "- B8 Libramiento San Ignacio";
                    return conexion;
                    break;

                case 9:
                    conexion = "- B9 Nazas";
                    return conexion;
                    break;
                case 10:
                    conexion = "- B10 Zarco";
                    return conexion;
                    break;

                case 11:
                    conexion = "- B11 Abastos Mayoreo";
                    return conexion;
                    break;
                case 12:
                    conexion = "- B12 Abastos Menudeo";
                    return conexion;
                    break;

                case 13:
                    conexion = "- B13 Miguel Aleman";
                    return conexion;
                    break;
                case 15:
                    conexion = "- B15 Triana";
                    return conexion;
                    break;

                case 16:
                    conexion = "- B16 Domingo Arrieta";
                    return conexion;
                    break;
                case 17:
                    conexion = "- B17 Centro Camionero Torreon";
                    return conexion;
                    break;

                case 18:
                    conexion = "- B18 Centro Camionero Menudeo";
                    return conexion;
                    break;
                default:
                    conexion = "NO SE RECONOCE SU CONFIGURACION";
                    return conexion;
                    break;

            }
        }
        public static string GetSetting(string appName, string section, string key, string sDefault)
        {
            // Los datos de VB se guardan en:
            // HKEY_CURRENT_USER\Software\VB and VBA Program Settings
            RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"Software\VB and VBA Program Settings\" +
                                                              appName + "\\" + section);
            string s = sDefault;
            if (rk != null)
            {
                s = (string)rk.GetValue(key);
            }
            return s;
        }

    }
    
}
