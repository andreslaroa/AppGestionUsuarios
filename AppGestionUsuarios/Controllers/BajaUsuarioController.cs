using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

[Authorize]
public class BajaUsuarioController : Controller
{
    [HttpGet]
    public IActionResult BajaUsuario()
    {
        try
        {
            int pageSize = 1000;
            using (DirectoryEntry entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = "(objectClass=user)";
                searcher.PageSize = pageSize;
                searcher.PropertiesToLoad.Add("displayName");
                searcher.PropertiesToLoad.Add("sAMAccountName");
                searcher.SearchScope = SearchScope.Subtree;

                List<string> usuarios = new List<string>();
                foreach (SearchResult result in searcher.FindAll())
                {
                    if (result.Properties.Contains("displayName") && result.Properties.Contains("sAMAccountName"))
                    {
                        string displayName = result.Properties["displayName"][0].ToString();
                        string samAccountName = result.Properties["sAMAccountName"][0].ToString();
                        usuarios.Add($"{displayName} ({samAccountName})");
                    }
                }

                ViewBag.Usuarios = usuarios.OrderBy(u => u).ToList();
            }
        }
        catch (Exception ex)
        {
            ViewBag.Usuarios = new List<string>();
            Console.WriteLine($"Error al cargar los usuarios: {ex.Message}");
        }

        return View();
    }

    [HttpPost]
    public IActionResult BajaUsuario([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username"))
            return Json(new { success = false, message = "Usuario no especificado." });

        string input = requestData["username"];
        string username = ExtractUsername(input); // Extraemos solo el sAMAccountName

        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });

        try
        {
            using (var context = new PrincipalContext(ContextType.Domain, "aytosa.inet"))
            using (var user = UserPrincipal.FindByIdentity(context, username))
            {
                if (user == null)
                    return Json(new { success = false, message = "Usuario no encontrado en Active Directory." });

                using (var userEntry = (DirectoryEntry)user.GetUnderlyingObject())
                {
                    string userDN = userEntry.Properties["distinguishedName"].Value.ToString();

                    // 1. Eliminar al usuario de todos los grupos
                    if (userEntry.Properties.Contains("memberOf"))
                    {
                        List<string> grupos = new List<string>();
                        foreach (var groupDN in userEntry.Properties["memberOf"])
                        {
                            string groupCN = ExtractCNFromDN(groupDN.ToString());
                            grupos.Add(groupCN);
                        }
                        foreach (string groupCN in grupos)
                        {
                            DirectoryEntry groupEntry = FindGroupByName(groupCN);
                            if (groupEntry != null)
                            {
                                try
                                {
                                    if (groupEntry.Properties["member"].Contains(userDN))
                                    {
                                        groupEntry.Properties["member"].Remove(userDN);
                                        groupEntry.CommitChanges();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error al eliminar usuario del grupo {groupCN}: {ex.Message}");
                                }
                            }
                        }
                    }

                    // 2. Eliminar la carpeta personal del usuario (si existe)
                    try
                    {
                        string userFolderPath = $"\\\\fs1.aytosa.inet\\home\\{username}";
                        if (Directory.Exists(userFolderPath))
                        {
                            Directory.Delete(userFolderPath, true);
                            Console.WriteLine($"Carpeta de usuario {username} eliminada correctamente.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al eliminar carpeta personal del usuario {username}: {ex.Message}");
                    }

                    // 3. Eliminar el usuario de Active Directory a través de su contenedor
                    try
                    {
                        DirectoryEntry parent = userEntry.Parent;
                        parent.Children.Remove(userEntry);
                        parent.CommitChanges();
                    }
                    catch (Exception ex)
                    {
                        return Json(new { success = false, message = $"Error al eliminar el usuario en AD: {ex.Message}" });
                    }
                }
            }
            return Json(new { success = true, message = "Usuario eliminado correctamente." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error general: {ex.Message}" });
        }
    }

    private string ExtractUsername(string input)
    {
        int startIndex = input.LastIndexOf('(');
        int endIndex = input.LastIndexOf(')');
        if (startIndex >= 0 && endIndex > startIndex)
        {
            return input.Substring(startIndex + 1, endIndex - startIndex - 1).Trim();
        }
        return null;
    }

    private string ExtractCNFromDN(string dn)
    {
        if (!string.IsNullOrEmpty(dn))
        {
            int start = dn.IndexOf("CN=");
            if (start >= 0)
            {
                int end = dn.IndexOf(",", start);
                return end > start
                    ? dn.Substring(start + 3, end - start - 3)
                    : dn.Substring(start + 3);
            }
        }
        return "";
    }

    private DirectoryEntry FindGroupByName(string groupName)
    {
        try
        {
            string domainPath = "LDAP://DC=aytosa,DC=inet";
            using (DirectoryEntry rootEntry = new DirectoryEntry(domainPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
                {
                    searcher.Filter = $"(&(objectClass=group)(cn={groupName}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");

                    SearchResult result = searcher.FindOne();
                    return result?.GetDirectoryEntry();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al buscar el grupo '{groupName}': {ex.Message}");
        }

        return null;
    }
}
