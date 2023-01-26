using PruebaOhPregunta2.Conexion;

namespace PruebaOhPregunta2
{
    class program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Ingrese el correo remitente");
            var email = Console.ReadLine();
            Console.WriteLine("Ingrese la clave del correo");
            var clave = Console.ReadLine();
            Console.WriteLine("Ingrese la fecha de inicio");
            var fechaInicio = Console.ReadLine();
            Console.WriteLine("Ingrese la fecha final");
            var fechaFinal = Console.ReadLine();
            
            ConexionBD conexion = new ConexionBD();

            conexion.listarSQL();

            EnvioCorreo correo = new EnvioCorreo();
            string body = @"<h1>Este es el body</h1>";
            correo.envioCorreo("franco.paredes@oechsle.pe", "Reporte Empleados - Examen Técnico Oechsle", body, email, clave);

            Console.Write($"{Environment.NewLine}Reporte de empleados desde {fechaInicio:d} hasta {fechaFinal:d}");
        }
    }
}
