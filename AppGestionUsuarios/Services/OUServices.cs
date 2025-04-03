using System.DirectoryServices;

public class OUService
{
    public OUService()
    {
    }

    public List<string> GetOUPrincipales()
    {
        // Este método ya no se usa directamente en AltaUsuarioController,
        // pero lo dejamos por si se necesita en otro lugar
        return new List<string>();
    }

    public List<string> GetOUSecundarias(string selectedOU)
    {
        try
        {
            var ouSecundarias = new List<string>();
            using (var rootEntry = new DirectoryEntry($"LDAP://OU=Usuarios y Grupos,OU={selectedOU},OU=AREAS,DC=aytosa,DC=inet"))
            {
                foreach (DirectoryEntry child in rootEntry.Children)
                {
                    if (child.SchemaClassName == "organizationalUnit")
                    {
                        string ouName = child.Properties["ou"].Value?.ToString();
                        if (!string.IsNullOrEmpty(ouName))
                        {
                            ouSecundarias.Add(ouName);
                        }
                    }
                }
            }
            ouSecundarias.Sort();
            return ouSecundarias;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener las OUs secundarias para {selectedOU}: {ex.Message}", ex);
        }
    }

    public List<string> GetDepartamentos(string selectedOU)
    {
        try
        {
            // Aquí podrías buscar departamentos en el AD, pero por ahora devolvemos una lista predeterminada
            return new List<string> { "Departamento1", "Departamento2", "Departamento3" };
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener los departamentos para {selectedOU}: {ex.Message}", ex);
        }
    }

    public List<string> GetLugarEnvio(string departamento)
    {
        try
        {
            // Aquí podrías buscar lugares de envío en el AD, pero por ahora devolvemos una lista predeterminada
            return new List<string> { "Lugar1", "Lugar2", "Lugar3" };
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener los lugares de envío para {departamento}: {ex.Message}", ex);
        }
    }

    public List<string> GetPortalEmpleado()
    {
        try
        {
            return new List<string> { "GA_R_PORTALDELEMPLEADO" };
        }
        catch (Exception ex)
        {
            throw new Exception("Error al obtener los portales del empleado: " + ex.Message, ex);
        }
    }

    public List<string> GetCuota()
    {
        try
        {
            return new List<string> { "500 MB", "1 GB", "2 GB" };
        }
        catch (Exception ex)
        {
            throw new Exception("Error al obtener las cuotas: " + ex.Message, ex);
        }
    }
}
