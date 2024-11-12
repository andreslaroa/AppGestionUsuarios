using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

public class OUService
{
    private readonly string _filePath;

    public OUService(string filePath)
    {
        _filePath = filePath;
    }

    public List<string> GetOUPrincipales()
    {
        var ouPrincipales = new List<string>();

        // Configurar la licencia de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            // Leer la primera hoja
            var worksheet = package.Workbook.Worksheets[0]; // Primera hoja
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Asume que la primera fila es el encabezado
            {
                string ouPrincipal = worksheet.Cells[row, 1].Text; // Columna 1
                if (!string.IsNullOrEmpty(ouPrincipal) && !ouPrincipales.Contains(ouPrincipal))
                {
                    ouPrincipales.Add(ouPrincipal);
                }
            }
        }

        return ouPrincipales;
    }

    public Dictionary<string, List<string>> GetOUSecundarias()
    {
        var ouSecundarias = new Dictionary<string, List<string>>();

        // Configurar la licencia de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            // Leer la segunda hoja
            var worksheet = package.Workbook.Worksheets[1]; // Segunda hoja
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Asume que la primera fila es el encabezado
            {
                string ouPrincipal = worksheet.Cells[row, 1].Text; // Columna 1
                string ouSecundaria = worksheet.Cells[row, 2].Text; // Columna 2

                if (!string.IsNullOrEmpty(ouPrincipal) && !string.IsNullOrEmpty(ouSecundaria))
                {
                    if (!ouSecundarias.ContainsKey(ouPrincipal))
                    {
                        ouSecundarias[ouPrincipal] = new List<string>();
                    }
                    ouSecundarias[ouPrincipal].Add(ouSecundaria);
                }
            }
        }

        return ouSecundarias;
    }
}
