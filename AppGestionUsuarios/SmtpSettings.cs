namespace AppGestionUsuarios.Notificaciones
{
    public class SmtpSettings
    {
        public string Server { get; set; }
        public int Port { get; set; }
        public string FromEmail { get; set; }
        public string AdminEmail { get; set; }
    }
}
