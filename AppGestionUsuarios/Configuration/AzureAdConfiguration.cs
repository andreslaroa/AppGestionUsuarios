namespace AppGestionUsuarios.Configuration
{
    
    /// Configuración para la conexión con Azure AD (Microsoft Graph)
    
    public class AzureAdSettings
    {
        
        /// Tenant (directorio) de Azure AD
        
        public string TenantId { get; set; } = string.Empty;

        
        /// Application (Client) Id de la app registrada en Azure AD
        
        public string ClientId { get; set; } = string.Empty;

        
        /// Secreto de cliente generado en Azure AD
        
        public string ClientSecret { get; set; } = string.Empty;
    }

    
    /// Configuración para la sincronización de Azure AD Connect (Delta Sync)
    
    public class AzureAdSyncSettings
    {
        
        /// Nombre o dirección del servidor donde corre Azure AD Connect (WinRM)
        
        public string SyncServer { get; set; } = string.Empty;

        
        /// Puerto WinRM (5985 HTTP, 5986 HTTPS)
        
        public int WsManPort { get; set; } = 5985;

        
        /// Mecanismo de autenticación para WinRM (Default, Negotiate, Kerberos...)
        
        public string Authentication { get; set; } = "Default";

        
        /// Retraso inicial en segundos antes de empezar a pollear (por defecto 30s)
        
        public int InitialDelaySeconds { get; set; } = 30;

        
        /// Intervalo en segundos entre cada poll (por defecto 10s)
        
        public int PollIntervalSeconds { get; set; } = 10;

        
        /// Número máximo de polls después del delay inicial (por defecto 6)
        
        public int MaxPollAttempts { get; set; } = 6;
    }
}
