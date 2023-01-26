using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using Aspose.Cells;

namespace PruebaOhPregunta2.Conexion
{
    internal class ConexionBD
    {
        private string Cnn = @"Server=mssql-107024-0.cloudclusters.net,10243;Initial Catalog=pruebaOh;User ID=ceiber;Password=TECINF2k10$;";
        private SqlConnection connection;
        private SqlCommand command;

        private SqlDataAdapter dataAdapter;
        private DataTable dataTable;
        private DataSet dataSet;

        private void Conexion()
        {
            connection = new SqlConnection(Cnn);
        }

        public ConexionBD() //Constructor
        {
            Conexion();
        }

        public DataTable listarSQL()
        {
            try
            {
                string sql = "SELECT * From Empleados";
                connection.Open();
                dataAdapter = new SqlDataAdapter(sql, connection);
                dataSet = new DataSet();
                dataAdapter.Fill(dataSet, "Empleados");
                dataTable = new DataTable();
                dataTable = dataSet.Tables["Empleados"];

                Workbook book = new Workbook();
                Worksheet sheet = book.Worksheets[0];

                ImportTableOptions importOptions = new ImportTableOptions();
                importOptions.ColumnIndexes = new int[] {0,1 };
                importOptions.IsFieldNameShown = true;
                sheet.Cells.ImportData(dataTable, 0, 0, importOptions);
                string path = Directory.GetCurrentDirectory() + "\\Reporte.xlsx";
                book.Save(path);

            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                connection.Close();
            }
            return dataTable;
        }
    }
}
