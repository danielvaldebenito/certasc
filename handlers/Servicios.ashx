<%@ WebHandler Language="C#" Class="Servicios" %>

using System;
using System.Web;
using System.Linq;
using System.Linq.Dynamic;
using System.Web.Script.Serialization;
using System.Globalization;
public class Servicios : IHttpHandler
{

    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
    (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    public void ProcessRequest(HttpContext context)
    {
        log4net.Config.DOMConfigurator.Configure();
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();
        string sidx, sord;
        int page, rows;

        switch (action)
        {
            case "grid":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var cliente = post.Request["cliente"];
                var desde = post.Request["desde"];
                var hasta = post.Request["hasta"];
                data = GetDataGrid(sidx, sord, page, rows, cliente, desde, hasta);
                break;
            case "gridInspecciones":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var service = int.Parse(post.Request["serviceId"]);
                data = GetDataInspecciones(sidx, sord, page, rows, service);
                break;
            case "addService":
                var cliente1 = int.Parse(post.Request["cliente"]);
                var it = post.Request["it"];
                data = AddService(cliente1, it);
                break;
            case "addInspeccion":

                data = AddInspeccion(post);
                break;
            case "editInspeccion":

                data = EditInspeccion(post);
                break;
            case "addClient":
                data = AddClient(post);
                break;
            default:
                context.Response.Write("No se ha seleccionado ninguna opción");
                break;
        }
        if (data != null)
        {
            var json = serializer.Serialize(data);
            context.Response.ContentType = "json";
            context.Response.Write(json);
            context.Response.Flush();

        }
    }
    private static object AddClient(HttpContext post)
    {
        try
        {
            var rut = post.Request["rut"];
            var nombre = post.Request["nombre"];
            var direccion = post.Request["direccion"];
            var telefono = post.Request["telefono"];
            var email = post.Request["email"];
            var giro = post.Request["giro"];
            using (var db = new CertelEntities())
            {
                var existClient = db.Cliente
                                    .Where(w => w.Rut == rut)
                                    .FirstOrDefault();
                if (existClient == null)
                {
                    var client = new Cliente
                    {
                        Rut = rut,
                        Nombre = nombre,
                        Direccion = direccion,
                        Habilitado = true,
                        EmailContacto = email,
                        TelefonoContacto = telefono,
                        Giro = giro
                    };
                    db.Cliente.Add(client);
                    db.SaveChanges();
                    return new
                    {
                        done = true,
                        message = "Cliente ingresado Correctamente",
                        id = client.ID

                    };
                }
                else
                {
                    return new
                    {
                        done = false,
                        message = "Ya existe un cliente con ese rut: (" + existClient.Nombre + ").",
                        code = 0
                    };
                }
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Servicios/AddClient", ex);
            return new
            {
                done = false,
                message = ex.ToString(),
                code = 0
            };
        }
    }
    private static object AddInspeccion(HttpContext post)
    {
        try
        {
            var ubicacion = post.Request["ubicacion"];
            var fechaInstalacion = post.Request["fechaInstalacion"];
            var fechaInspeccion = post.Request["fechaInspeccion"];
            var aparato = int.Parse(post.Request["aparato"]);
            var funcionamiento = int.Parse(post.Request["funcionamiento"]);
            var destino = int.Parse(post.Request["destino"]);
            var permiso = post.Request["permiso"];
            var recepcion = post.Request["recepcion"];
            var altura = int.Parse(post.Request["altura"]);
            var ingeniero = post.Request["ingeniero"];
            var servicio = int.Parse(post.Request["servicio"]);
            var itServicio = post.Request["itServicio"];
            var nombreProyecto = post.Request["nombre"];
            var numero = post.Request["numero"];
            var nombreEdificio = post.Request["edificio"];
            var fec = post.Request["fec"];
            var fvc = post.Request["fvc"];
            using (var db = new CertelEntities())
            {
                var service = db.Servicio
                    .Where(w => servicio == 0 ? w.IT == itServicio : w.ID == servicio)
                                .Select(s => new { Id = s.ID, It = s.IT, Count = s.Inspeccion.Count() })
                                .FirstOrDefault();

                if (service == null)
                {
                    return new
                    {
                        done = false,
                        message = "IT de servicio no existe",
                        code = 1
                    };
                }

                var newIt = GetNewIt(service.It);
                var inspeccion = new Inspeccion
                {
                    Ubicacion = ubicacion,
                    FechaCreacion = DateTime.Now,
                    AparatoID = aparato,
                    TipoFuncionamientoID = funcionamiento,
                    DestinoProyectoID = destino,
                    PermisoEdificacion = permiso,
                    RecepcionMunicipal = recepcion,
                    AlturaPisos = altura,
                    Ingeniero = ingeniero,
                    ServicioID = service.Id,
                    Fase = 1,
                    IT = newIt,
                    EstadoID = 1,
                    Numero = numero,
                    NombreEdificio = nombreEdificio,
                    NombreProyecto = nombreProyecto
                };
                if (fechaInspeccion != string.Empty)
                {
                    inspeccion.FechaInspeccion = DateTime.ParseExact(fechaInspeccion, "dd-MM-yyyy", null);
                    inspeccion.FechaEntrega = DateTime.Now.AddDays(5);
                }
                if (fechaInstalacion != string.Empty)
                    inspeccion.FechaInstalacion = DateTime.ParseExact(fechaInstalacion, "dd-MM-yyyy", null);

                if(fec != string.Empty)
                {
                    inspeccion.FechaEmisionCertificado = DateTime.ParseExact(fec, "dd-MM-yyyy", null);
                }
                if(fvc != string.Empty)
                {
                    inspeccion.FechaVencimientoCertificado = DateTime.ParseExact(fvc, "dd-MM-yyyy", null);
                }
                db.Inspeccion.Add(inspeccion);
                db.SaveChanges();
                return new
                {
                    done = true,
                    message = "Inspección creada exitosamente",
                    id = inspeccion.ID
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Servicios/AddInspeccion", ex);
            return new
            {
                done = false,
                message = ex.ToString(),
                code = 0
            };
        }
    }

    private static string GetNewIt(string it)
    {
        using (var db = new CertelEntities())
        {
            var service = db.Servicio.Where(w => w.IT == it).FirstOrDefault();
            var inspecciones = service.Inspeccion;
            var digit = 0;
            while (true)
            {
                digit++;
                var existsInspeccion = inspecciones.Any(a => int.Parse(a.IT.Split('/')[1]) == digit);
                if (!existsInspeccion) break;
            }
            return string.Format("{0}/{1}", service.IT, digit);
        }



    }
    // Edit
    private static object EditInspeccion(HttpContext post)
    {
        try
        {
            var ubicacion = post.Request["ubicacion"];
            var fechaInstalacion = post.Request["fechaInstalacion"];
            var fechaInspeccion = post.Request["fechaInspeccion"];
            var aparato = int.Parse(post.Request["aparato"]);
            var funcionamiento = int.Parse(post.Request["funcionamiento"]);
            var destino = int.Parse(post.Request["destino"]);
            var permiso = post.Request["permiso"];
            var recepcion = post.Request["recepcion"];
            var altura = int.Parse(post.Request["altura"]);
            var ingeniero = post.Request["ingeniero"];
            var nombreProyecto = post.Request["nombre"];
            var numero = post.Request["numero"];
            var nombreEdificio = post.Request["edificio"];
            var fec = post.Request["fec"];
            var fvc = post.Request["fvc"];
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var inspeccion = db.Inspeccion
                                    .Find(id);
                if (inspeccion == null)
                {
                    return new
                    {
                        done = false,
                        message = "Inspección no existe (¿?)"
                    };
                }
                inspeccion.Ubicacion = ubicacion;
                inspeccion.AparatoID = aparato;
                inspeccion.TipoFuncionamientoID = funcionamiento;
                inspeccion.DestinoProyectoID = destino;
                inspeccion.PermisoEdificacion = permiso;
                inspeccion.RecepcionMunicipal = recepcion;
                inspeccion.AlturaPisos = altura;
                inspeccion.Ingeniero = ingeniero;
                inspeccion.Numero = numero;
                inspeccion.NombreProyecto = nombreProyecto;
                inspeccion.NombreEdificio = nombreEdificio;

                if (fechaInspeccion != string.Empty)
                {
                    inspeccion.FechaInspeccion = DateTime.ParseExact(fechaInspeccion, "dd-MM-yyyy", null);
                    inspeccion.FechaEntrega = DateTime.Now.AddDays(5);
                }

                if (fechaInstalacion != string.Empty)
                    inspeccion.FechaInstalacion = DateTime.ParseExact(fechaInstalacion, "dd-MM-yyyy", null);

                if(fec != string.Empty)
                {
                    inspeccion.FechaEmisionCertificado = DateTime.ParseExact(fec, "dd-MM-yyyy", null);
                }
                if(fvc != string.Empty)
                {
                    inspeccion.FechaVencimientoCertificado = DateTime.ParseExact(fvc, "dd-MM-yyyy", null);
                }
                db.SaveChanges();
                return new
                {
                    done = true,
                    message = "Inspección modificada exitosamente",
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Servicios/EditInspeccion", ex);
            return new
            {
                done = false,
                message = ex.ToString(),
                code = 0
            };
        }
    }

    private static object AddService(int cliente, string it)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var exists = db.Servicio
                                .Where(w => w.IT == it)
                                .FirstOrDefault();
                if (exists != null)
                    return new
                    {
                        done = false,
                        message = "Ya existe un servicio con el IT: " + it + ""
                    };

                var service = new Servicio
                {
                    ClienteID = cliente,
                    IT = it,
                    EstadoID = 1,
                    FechaCreacion = DateTime.Now,
                };
                db.Servicio.Add(service);
                db.SaveChanges();
                return new
                {
                    done = true,
                    message = "Servicio creado exitosamente",
                    id = service.ID,
                    it = service.IT
                };
            }
        }
        catch (Exception ex)
        {
            log.Error("Excepción Servicios/AddService", ex);
            return new
            {
                done = false,
                message = ex.ToString()
            };
        }
    }
    private static object GetDataGrid(string sidx, string sord, int page, int rows, string cliente, string desde, string hasta)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var dt_desde = desde == string.Empty ? new DateTime(2000, 01, 01) : DateTime.ParseExact(desde, "dd-MM-yyyy", null);
                var dt_hasta = hasta == string.Empty ? DateTime.Now : DateTime.ParseExact(hasta, "dd-MM-yyyy", null).AddDays(1);
                var list = db.Servicio
                                .Where(w => w.Cliente.Nombre.Contains(cliente))
                                .Where(w => w.FechaCreacion >= dt_desde)
                                .Where(w => w.FechaCreacion <= dt_hasta)
                                .ToList();
                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    .OrderBy(sidx + " " + sord)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);
                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result

                           .Select(x => new
                           {
                               Id = x.ID,
                               FechaCreacion = x.FechaCreacion.ToString("dd-MM-yyyy"),

                               Cliente = x.ClienteID == null ? string.Empty : x.Cliente.Nombre,
                               IT = x.IT,
                               EstadoId = x.EstadoID,
                               Estado = x.EstadoServicio.Descripcion
                           })
                           .ToList()
                };
                return grid;
            }

        }
        catch (Exception ex)
        {
            log.Error("ERROR AL CARGAR GRILLA DE SERVICIOS", ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object GetDataInspecciones(string sidx, string sord, int page, int rows, int service)
    {
        try
        {
            using (var db = new CertelEntities())
            {

                var list = db.Inspeccion
                                .Where(w => w.ServicioID == service)
                                .ToList();
                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    .OrderBy(sidx + " " + sord)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);
                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result

                           .Select(x => new
                           {
                               Id = x.ID,
                               IT = x.IT,
                               FechaCreacion = x.FechaCreacion.ToString("dd-MM-yyyy"),
                               FechaInspeccion = x.FechaInspeccion == null ? string.Empty : ((DateTime)x.FechaInspeccion).ToString("dd-MM-yyyy"),
                               Aparato = x.Aparato.Nombre,
                               Funcionamiento = x.TipoFuncionamientoAparato.Descripcion,
                               Fase = x.Fase,
                               EstadoId = x.EstadoID,
                               Estado = x.EstadoInspeccion.Descripcion
                           })
                           .ToList()
                };
                return grid;
            }

        }
        catch (Exception ex)
        {
            log.Error("ERROR AL CARGAR GRILLA DE INSPECCIONES", ex);
            return new { done = false, message = ex.ToString() };
        }
    }
    public bool IsReusable
    {
        get
        {
            return true;
        }
    }

}