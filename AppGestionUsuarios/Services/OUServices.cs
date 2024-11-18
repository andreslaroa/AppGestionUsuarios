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
            var worksheet = package.Workbook.Worksheets["OU_PRINCIPAL"]; // Primera hoja
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

    public List<string> GetOUSecundarias(string selectedOU)
    {
        var ouSecundarias = new List<string>();

        // Configurar la licencia de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets["OU_SECUNDARIA"]; // Hoja llamada "OU_SECUNDARIA"

            if (worksheet == null)
                return ouSecundarias; // Si la hoja no existe, retorna la lista vacía

            int columnCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;

            // Encontrar la columna correspondiente a la OU seleccionada (buscando en la primera fila)
            int targetColumnIndex = -1;
            for (int col = 1; col <= columnCount; col++)
            {
                string header = worksheet.Cells[1, col].Text; // Primera fila (encabezado)
                if (header.Equals(selectedOU, System.StringComparison.OrdinalIgnoreCase))
                {
                    targetColumnIndex = col;
                    break;
                }
            }

            if (targetColumnIndex == -1)
                return ouSecundarias; // No se encontró la OU seleccionada en la primera fila

            // Recorrer la columna y agregar celdas con contenido a la lista
            for (int row = 2; row <= rowCount; row++) // Comienza en la fila 2 para omitir el encabezado
            {
                string ouSecundaria = worksheet.Cells[row, targetColumnIndex].Text;
                if (!string.IsNullOrEmpty(ouSecundaria))
                {
                    ouSecundarias.Add(ouSecundaria);
                }
            }
        }

        return ouSecundarias;
    }


    public List<string> GetDepartamentos(string selectedOU)
    {
        var departamentos = new List<string>();

        // Configurar la licencia de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets["DEPARTAMENTO"]; // Hoja llamada "DEPARTAMENTO"

            if (worksheet == null)
                return departamentos; // Si la hoja no existe, retorna la lista vacía

            int columnCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;

            // Encontrar la columna correspondiente al departamento seleccionada (buscando en la primera fila)
            int targetColumnIndex = -1;
            for (int col = 1; col <= columnCount; col++)
            {
                string header = worksheet.Cells[1, col].Text; 
                if (header.Equals(selectedOU, System.StringComparison.OrdinalIgnoreCase))
                {
                    targetColumnIndex = col;
                    break;
                }
            }

            if (targetColumnIndex == -1)
                return departamentos; // No se encontró la OU seleccionada en la primera fila

            // Recorrer la columna y agregar celdas con contenido a la lista
            for (int row = 2; row <= rowCount; row++) // Comienza en la fila 2 para omitir el encabezado
            {
                string departamento = worksheet.Cells[row, targetColumnIndex].Text;
                if (!string.IsNullOrEmpty(departamento))
                {
                    departamentos.Add(departamento);
                }
            }
        }

        return departamentos;
    }


}
