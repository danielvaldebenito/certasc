<%@ WebHandler Language="C#" Class="ExportExcel" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using System.Data;
using log4net;
using System.Web.SessionState;
using System.Globalization;
public class ExportExcel : IHttpHandler {

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
    (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest (HttpContext context) {
        log4net.Config.DOMConfigurator.Configure();
        var post = HttpContext.Current;
        var accion = post.Request["1"].Trim();
        object data = new object();
        switch (accion)
        {
            case "cotizaciones":
                ExportCotizaciones(post);
                break;
            case "clientes":
                ExportClientes(post);
                break;
            case "cotizacionesCliente":
                ExportCotizacionesCliente(post);
                break;
        }
    }

    private static void ExportCotizacionesCliente (HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var cliente = db.Cliente.Find(id);
                var list = db.Cotizacion
                                .Where(w => w.ClienteId == id)
                                .AsEnumerable()
                                .Select(s => new Cot
                                {
                                    Id = s.Id,
                                    IT = s.IT.ToString(),
                                    Cliente = s.Cliente.Nombre,
                                    NombreContacto = s.Cliente.NombreContacto,
                                    FechaCreacion = s.FechaCreacion.Value.ToString("dd-MM-yyyy HH:mm:ss"),
                                    Estado = s.EstadoCotizacion.Descripcion,
                                    FechaDoc = s.FechaDoc.Value.ToString("dd-MM-yyyy"),
                                    Responsable = s.Usuario1.Nombre,
                                    Moneda = s.Moneda.Descripcion,
                                    Valor = s.Valor == null ? 0 : (double)s.Valor,
                                    Descuento =  s.Descuento == null ? 0 :  (double)s.Descuento,
                                    Recargo =  s.Recargo == null ? 0 : (double)s.Recargo,
                                    Total = s.Total == null ? 0 :  (double)s.Total,
                                    Observacion = s.Observacion,
                                    TotalServicios = s.Servicio.Sum(ss => ss.Inspeccion.Count),
                                    TotalCalifica = s.Servicio.Sum(w => w.Inspeccion.Where(ww => ww.Calificacion == 1).Count()),
                                    TotalCalificaConObservaciones = s.Servicio.Sum(w => w.Inspeccion.Where(ww => ww.Calificacion == 2).Count()),
                                    TotalNoCalifica = s.Servicio.Sum(w => w.Inspeccion.Where(ww => ww.Calificacion == 0).Count()),
                                    
                                })
                                .ToList();

                var filename = "COTIZACIONES CLIENTES " + cliente.Nombre + " " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + ".xls";
                Export(filename, list);
            }
        }
        catch (Exception ex)
        {
            post.Response.Write(ex.ToString());
            return;
        }
    }

    private static void ExportClientes (HttpContext post)
    {
        try
        {
            var name = post.Request["name"];
            var rut = post.Request["rut"];
            using (var db = new CertelEntities())
            {
                var list = db.Cliente
                                .AsEnumerable()
                                .Select(s => new Cli
                                {
                                    Rut = s.Rut,
                                    Nombre = s.Nombre,
                                    Giro = s.Giro,
                                    Direccion = s.Direccion,
                                    Telefono = s.TelefonoContacto,
                                    Email = s.EmailContacto,
                                    Habilitado = s.Habilitado ? "SÍ" : "NO",
                                    CotizacionesCreadas = s.Cotizacion.Count(a => a.EstadoId == 1),
                                    CotizacionesEnEvaluacion = s.Cotizacion.Count(c => c.EstadoId == 2),
                                    CotizacionesEnEjecucion = s.Cotizacion.Count(c => c.EstadoId == 3),
                                    CotizacionesCanceladas = s.Cotizacion.Count(c => c.EstadoId == 4),
                                    CotizacionesFinalizadas = s.Cotizacion.Count(c => c.EstadoId == 5)
                                })
                                .ToList();

                var filename = "CLIENTES AL " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + ".xls";
                Export(filename, list);
            }
        }
        catch (Exception ex)
        {
            post.Response.Write(ex.ToString());
            return;
        }
    }
    private static void ExportCotizaciones (HttpContext post)
    {
        try
        {
            var cliente = post.Request["cliente"];
            var estado = int.Parse(post.Request["estado"]);
            var inicio = post.Request["inicio"];
            var fin = post.Request["fin"];
            var f1 = inicio == string.Empty ? new DateTime(2017, 1, 1) : DateTime.ParseExact(inicio, "dd-MM-yyyy", null);
            var f2 = fin == string.Empty ? DateTime.Now : DateTime.ParseExact(fin, "dd-MM-yyyy", null);

            using (var db = new CertelEntities())
            {
                var list = db.Cotizacion
                        .Where(w => w.Cliente.Nombre.Contains(cliente) || w.Cliente.Rut.Contains(cliente))
                        .Where(w => estado == 0 ? true : w.EstadoId == estado)
                        .Where(w => w.FechaDoc >= f1 && w.FechaDoc <= f2)
                        .AsEnumerable()
                        .Select(s => new Cot {
                            Id = s.Id,
                            IT = s.IT.ToString(),
                            Cliente = s.Cliente.Nombre,
                            NombreContacto = s.Cliente.NombreContacto,
                            FechaCreacion = s.FechaCreacion.Value.ToString("dd-MM-yyyy HH:mm:ss"),
                            Estado = s.EstadoCotizacion.Descripcion,
                            FechaDoc = s.FechaDoc.Value.ToString("dd-MM-yyyy"),
                            Responsable = s.Usuario1.Nombre,
                            Moneda = s.Moneda.Descripcion,
                            Valor = s.Valor == null ? 0 : (double)s.Valor,
                            Descuento =  s.Descuento == null ? 0 :  (double)s.Descuento,
                            Recargo =  s.Recargo == null ? 0 : (double)s.Recargo,
                            Total = s.Total == null ? 0 :  (double)s.Total,
                            Observacion = s.Observacion
                        })
                        .ToList()
                        ;
                var filename = "COTIZACIONES AL " + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + ".xls";
                Export(filename, list);
            }
        }
        catch(Exception ex)
        {
            post.Response.Write(ex.ToString());
            return;
        }
    }
    public struct Cot
    {
        public int Id { get; set; }
        public string IT { get; set; }
        public string Cliente { get; set; }
        public string FechaDoc { get; set; }
        public string NombreContacto { get; set; }
        public string Responsable { get; set; }
        public string Estado { get; set; }
        public string Moneda { get; set; }
        public string Observacion { get; set; }
        public string FechaCreacion { get; set; }
        public double Valor { get; set; }
        public double Descuento { get; set; }
        public double Recargo { get; set; }
        public double Total { get; set; }
        public int TotalServicios { get; set; }
        public int TotalCalifica { get; set; }
        public int TotalCalificaConObservaciones { get; set; }
        public int TotalNoCalifica { get; set; }
    }
    public struct Cli
    {
        public string Rut { get; set; }
        public string Nombre { get; set; }
        public string Giro { get; set; }
        public string Direccion { get; set; }
        public string Telefono { get; set; }
        public string Email { get; set; }
        public string Habilitado { get; set; }
        public int CotizacionesCreadas { get; set; }
        public int CotizacionesEnEvaluacion { get; set; }
        public int CotizacionesEnEjecucion { get; set; }
        public int CotizacionesCanceladas { get; set; }
        public int CotizacionesFinalizadas { get; set; }

    }
    public static void Export(string fileName, List<Cot> list)
    {
        try
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.AddHeader(
            "content-disposition", string.Format("attachment; filename={0}", fileName));
            HttpContext.Current.Response.ContentType = "Application/x-msexcel";
            HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1252");
            HttpContext.Current.Response.Charset = "utf-8";
            string tab = "";
            int cont = 0;
            foreach (var fila in list)
            {
                tab = "";
                System.Reflection.PropertyInfo[] props = fila.GetType().GetProperties();
                if (cont == 0)
                {
                    foreach (System.Reflection.PropertyInfo p in props)
                    {
                        HttpContext.Current.Response.Write(tab + p.Name);
                        tab = "\t";
                    }
                    cont++;
                }
                else
                    break;
            }
            HttpContext.Current.Response.Write("\n");
            foreach (var fila in list)
            {
                tab = "";
                System.Reflection.PropertyInfo[] props = fila.GetType().GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {

                    HttpContext.Current.Response.Write(tab + p.GetValue(fila, null));
                    tab = "\t";
                }
                HttpContext.Current.Response.Write("\n");
            }
            HttpContext.Current.Response.OutputStream.Close();

        }
        catch (Exception e)
        {

        }
    }


    public static void Export(string fileName, List<Cli> list)
    {
        try
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.AddHeader(
            "content-disposition", string.Format("attachment; filename={0}", fileName));
            HttpContext.Current.Response.ContentType = "Application/x-msexcel";
            HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1252");
            HttpContext.Current.Response.Charset = "utf-8";
            string tab = "";
            int cont = 0;
            foreach (var fila in list)
            {
                tab = "";
                System.Reflection.PropertyInfo[] props = fila.GetType().GetProperties();
                if (cont == 0)
                {
                    foreach (System.Reflection.PropertyInfo p in props)
                    {
                        HttpContext.Current.Response.Write(tab + p.Name);
                        tab = "\t";
                    }
                    cont++;
                }
                else
                    break;
            }
            HttpContext.Current.Response.Write("\n");
            foreach (var fila in list)
            {
                tab = "";
                System.Reflection.PropertyInfo[] props = fila.GetType().GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {

                    HttpContext.Current.Response.Write(tab + p.GetValue(fila, null));
                    tab = "\t";
                }
                HttpContext.Current.Response.Write("\n");
            }
            HttpContext.Current.Response.OutputStream.Close();

        }
        catch (Exception e)
        {

        }
    }



    public bool IsReusable {
        get {
            return true;
        }
    }

}