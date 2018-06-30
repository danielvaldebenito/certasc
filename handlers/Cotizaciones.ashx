<%@ WebHandler Language="C#" Class="Cotizaciones" %>

using System;
using System.Web;
using System.Linq;
using System.Web.Script.Serialization;
using System.Linq.Dynamic;
using System.Web.SessionState;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Net.Mail;
using System.Globalization;
public class Cotizaciones : IHttpHandler, IRequiresSessionState
{
    private static readonly log4net.ILog log = log4net.LogManager.GetLogger
   (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    //public static ConfigVariables config { get; set; }
    public void ProcessRequest(HttpContext context)
    {
        log4net.Config.DOMConfigurator.Configure();
        var post = HttpContext.Current;
        var data = new object();
        var action = post.Request["1"];
        var serializer = new JavaScriptSerializer();
        string sidx, sord;
        int page, rows;
        Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
        //config = new ConfigVariables();
        switch (action)
        {
            case "grid-clientes":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var name = post.Request["name"];
                var rut = post.Request["rut"];
                data = GetClientes(sidx, sord, page, rows, name, rut);
                break;
            case "grid-productos":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var name1 = post.Request["name"];
                data = GetProductos(sidx, sord, page, rows, name1);
                break;
            case "grid-cotizaciones":
                sidx = post.Request["sidx"];
                sord = post.Request["sord"];
                page = int.Parse(post.Request["page"]);
                rows = int.Parse(post.Request["rows"]);
                var cliente = post.Request["cliente"];
                var estado = int.Parse(post.Request["estado"]);
                var inicio = post.Request["inicio"];
                var fin = post.Request["fin"];
                var alertF2 = bool.Parse(post.Request["alertF2"]);
                data = GetCotizaciones(sidx, sord, page, rows, cliente, estado, inicio, fin, alertF2);
                break;
            case "save":
                data = Save(post);
                break;
            case "mail-cotizacion":
                data = Reenviar(post,  true);
                break;
            case "get-pdf":
                data = GetPDF(post);
                break;
            case "get-global":
                data = GetGlobal(post);
                break;
            case "get-files":
                var type = post.Request["type"];
                data = GetFiles(type);
                break;
            case "getClientByRut":
                var rut1 = post.Request["rut"];
                data = GetClientByRut(rut1);
                break;
            case "setPrice":
                data = GetPrice(post);
                break;
            case "getPriceLogisticCoste":
                data = GetPriceLogisticCoste(post);
                break;
            case "uploadFile":
                data = UploadFile(post);
                break;
            case "disableFile":
                data = DisableFile(post);
                break;
            case "getOne":
                data = GetOne(post);
                break;
            case "getServicesFromIT":
                data = GetServicesFromIT(post);
                break;
            case "reenviar":
                data = Reenviar(post, false);
                break;
            case "crearFase2":
                data = CrearFase2(post);
                break;
            case "mailPosVenta":
                data = MailPosVenta(post);
                break;
            case "sendAlertaFase2":
                data = SendAlertaFase2(post);
                break;
            case "completingSendAlertaFase2":
                data = CompletingSendAlertaFase2(post);
                break;
        }
        if (data != null)
        {
            var json = serializer.Serialize(data);

            context.Response.ContentType = "application/json";
            context.Response.Write(json);
            context.Response.Flush();

        }
    }
    private static object CompletingSendAlertaFase2(HttpContext post)
    {
        try
        {
            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var config = new ConfigVariables();
                config.SetData();
                var format = config.mail_alerta_fase_2;
                var cotizacion = db.Cotizacion.Find(id);
                if (cotizacion == null) return new { done = false, message = "Error: No existe cotizacion " + id };
                var item = cotizacion.ItemCotizacion.Select(s => s.Ubicacion).FirstOrDefault();
                var mail = string.Format(format,
                    cotizacion.Cliente.Nombre,
                    item ?? string.Empty,
                    cotizacion.FechaCompromisoFase2.Value.ToLongDateString(),
                    cotizacion.FechaCompromisoFase2.Value.AddDays(-7).ToLongDateString()
                );
                var direccion = cotizacion.Cliente.EmailContacto;
                return new { done = true, mail, direccion };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = "Ocurrió un error", ex = ex.ToString() };
        }
    }
    private static object SendAlertaFase2(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);
                var user = post.Request["user"];
                var to = post.Request["to"];
                var subject = post.Request["subject"];
                var mail = post.Request["mail"];
                var cotizacion = db.Cotizacion.Find(id);
                var usuario = new Encriptacion(user, false).newText;
                if (cotizacion == null) return new { done = false, message = "Cotización no existe" };
                var us = db.Usuario.Find(usuario);
                if (us == null) return new { done = false, message = "No se reconoce al usuario" };
                var from = us.Email;
                var pass = new Encriptacion(us.PassMail, false).newText;
                var name = us.Nombre + " " + us.Apellido;
                cotizacion.AlertaFase2Enviada = true;
                db.SaveChanges();
                SendMail(to, "", "", subject, from, pass, mail, cotizacion, from, name, us.Firma, false);
                return new { done = true, message = "Mensaje entregado exitosamente" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString(), message = "Ha ocurrido un error " };
        }
    }

    private static object MailPosVenta(HttpContext post)
    {
        try
        {
            var to = post.Request["to"];
            var asunto = post.Request["asunto"];
            var correo = post.Request["correo"];
            var encriptuser = post.Request["user"];
            var user = new Encriptacion(encriptuser, false).newText;

            var id = int.Parse(post.Request["id"]);
            using (var db = new CertelEntities())
            {
                var usuario = db.Usuario.Find(user);
                var cotizacion = db.Cotizacion.Find(id);
                if (cotizacion == null) return new { done = false, message = "Error: Cotización no existe" };
                var hito = new Hito
                {
                    CotizacionId = id,
                    Fecha = DateTime.Now,
                    TipoHitoId = 20,
                    Observacion = asunto
                };
                db.Hito.Add(hito);
                cotizacion.EstadoId = 5;
                db.SaveChanges();
                var from = usuario.Email;
                var pass = new Encriptacion(usuario.PassMail, false).newText;
                var name = usuario.Nombre + " " + usuario.Apellido;
                SendMail(to, string.Empty, string.Empty, asunto, from, pass, correo, cotizacion, from, name, usuario.Firma, false);
                return new { done = true, message = "Cotización finalizada exitosamente" };
            }

        }
        catch (Exception ex)
        {
            return new { done = false, message = "Ha ocurrido un error", ex = ex.ToString() };
        }
    }
    private static object CrearFase2(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);
                var date = post.Request["date"];
                var cotizacion = db.Cotizacion.Find(id);
                if (cotizacion == null)
                    return new { done = false, message = "Error: Cotización no existe" };
                if (cotizacion.CreadaFase2 == true)
                    return new { done = false, message = "Ya está creada la fase 2 para esta cotización" };
                var service = cotizacion.Servicio.FirstOrDefault();
                if (service == null)
                    return new { done = false, message = "No hay servicio asociado aún" };
                var inspecciones = service.Inspeccion.ToList();
                if (!inspecciones.Any())
                    return new { done = false, message = "No hay inspecciones asociadas a la cotización" };
                var calificacionesParaFase2 = new[] { 0, 2 };
                var count = 0;
                foreach(var inspeccion in inspecciones)
                {
                    if (inspeccion.Calificacion == null) continue;
                    if(calificacionesParaFase2.Contains((int)inspeccion.Calificacion) && inspeccion.Fase == 1 && !inspeccion.Inspeccion1.Any())
                    {
                        var fase2 = new Inspeccion
                        {
                            EstadoID = 1,
                            Fase = 2,
                            FechaCreacion = DateTime.Now,
                            InspeccionFase1 = inspeccion.ID,
                            AparatoID = inspeccion.AparatoID,
                            AlturaPisos = inspeccion.AlturaPisos,
                            Destinatario = inspeccion.Destinatario,
                            DestinoProyectoID = inspeccion.DestinoProyectoID,
                            FechaInstalacion = inspeccion.FechaInstalacion,
                            Ingeniero = inspeccion.Ingeniero,
                            IT = inspeccion.IT,
                            NombreEdificio = inspeccion.NombreEdificio,
                            NombreProyecto = inspeccion.NombreProyecto,
                            TipoInformeID = inspeccion.TipoInformeID,
                            TipoFuncionamientoID = inspeccion.TipoFuncionamientoID,
                            ServicioID = inspeccion.ServicioID,
                            Numero = inspeccion.Numero,
                            Ubicacion = inspeccion.Ubicacion,
                            PermisoEdificacion = inspeccion.PermisoEdificacion,
                            RecepcionMunicipal = inspeccion.RecepcionMunicipal,
                            FechaInspeccion = DateTime.ParseExact(date, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture),
                            CreaFaseSiguiente = false
                        };

                        db.Inspeccion.Add(fase2);
                        db.SaveChanges();
                        var inspeccionNormaOriginal = inspeccion.InspeccionNorma;
                        foreach (var i in inspeccionNormaOriginal)
                        {
                            var inspeccionNorma = new InspeccionNorma
                            {
                                InspeccionID = fase2.ID,
                                NormaID = i.NormaID
                            };
                            db.InspeccionNorma.Add(inspeccionNorma);
                        }
                        count++;
                    }
                    else
                    {
                        continue;
                    }
                }
                var hito = new Hito
                {
                    TipoHitoId = 19, // Creación Fase II
                    CotizacionId = id,
                    Fecha = DateTime.Now,
                    Observacion = "",
                    FechaCompromiso = DateTime.ParseExact(date, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture)
                };
                db.Hito.Add(hito);
                if(count > 0) {
                    cotizacion.CreadaFase2 = true;
                    cotizacion.FechaCompromisoFase2 = DateTime.ParseExact(date, "dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture);
                }

                db.SaveChanges();
                return new { done = true, message = count == 0 ? "Ninguno de los servicios cumple condiciones para crear la fase II" : "Fase II creada exitosamente con " + count + " inspecciones." };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = "Ha ocurrido un error", ex = ex.ToString() };
        }
    }
    private static object GetServicesFromIT(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var it = post.Request["it"];
                var rows = int.Parse(post.Request["rows"]);
                var sidx = post.Request["sidx"];
                var sord = post.Request["sord"];
                var page = int.Parse(post.Request["page"]);
                var services = db.Inspeccion
                                .Where(w => w.Servicio.IT == it)
                                .Select(i => new
                                {
                                    Id = i.ID,
                                    It = i.IT,
                                    Fase = i.Fase,
                                    Calificacion = i.Calificacion,
                                    Calificacion1 = i.Calificacion,
                                    Estado = i.EstadoInspeccion.Descripcion
                                });

                int pageSize = rows;
                int totalRecords = services.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = services
                                .OrderBy(sidx + " " + sord)
                                .Skip((page - 1) * pageSize)
                                .Take(pageSize);
                return new {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result.ToList()
                };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetOne(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var it = int.Parse(post.Request["it"]);
                var cotizacion = db.Cotizacion.Where(w => w.IT == it).FirstOrDefault();
                if (cotizacion == null)
                    return new { done = false, message = "No existe cotización con el IT ingresado" };
                return new
                {
                    done = true,
                    data = new
                    {
                        Cliente = new
                        {
                            Rut = cotizacion.Cliente.Rut,
                            Nombre = cotizacion.Cliente.Nombre,
                            Giro = cotizacion.Cliente.Giro,
                            Direccion = cotizacion.Cliente.Direccion,
                            NombreContacto = cotizacion.Cliente.NombreContacto,
                            EmailContacto = cotizacion.Cliente.EmailContacto,
                            TelefonoContacto = cotizacion.Cliente.TelefonoContacto
                        },
                        DatosGenerales = new
                        {
                            Id = cotizacion.Id,
                            It = cotizacion.IT,
                            FechaDoc = cotizacion.FechaDoc.HasValue ? cotizacion.FechaDoc.Value.ToString("dd-MM-yyyy") : string.Empty,
                            FuenteSolicitud = cotizacion.FuenteSolicitud,
                            FormaDePago = cotizacion.FormaDePagoId,
                            Moneda = cotizacion.MonedaId,
                            Vendedor = cotizacion.Vendedor,
                            Ciudad = cotizacion.CiudadId,
                            FechaValidez = cotizacion.FechaValidez.HasValue ? cotizacion.FechaValidez.Value.ToString("dd-MM-yyyy") : string.Empty,
                            Observacion = cotizacion.Observacion,
                            Nota = cotizacion.Nota
                        },
                        Items = cotizacion.ItemCotizacion
                            .Select(i => new
                            {
                                Cantidad = i.Cantidad,
                                Descripcion = i.DescripcionEditada,
                                Ubicacion = i.Ubicacion,
                                ValorUnitario = i.Unitario,
                                ModoDescuento = i.ModoDescuento,
                                ValorDescuento = i.ValorDescuento,
                                Producto = i.ProductoId,
                                Descuento = i.Descuento,
                                TipoElevador = i.AparatoID,
                                TipoElevadorName = i.AparatoID == null ? string.Empty : i.Aparato.Nombre,
                                TipoFuncionamientoName = i.TipoFuncionamientoID == null ? string.Empty : i.TipoFuncionamientoAparato.Descripcion,
                                Marca = i.MarcaID,
                                MarcaName = i.MarcaID == null ? string.Empty : i.Marca.Descripcion,
                                TipoFuncionamiento = i.TipoFuncionamientoID,
                                AnoInstalacion = i.InstalacionAno,
                                EmpresaInstaladora = i.EmpresaInstaladora,
                                Uso = i.Uso,
                                Altura = i.AlturaPisos
                            }).ToList(),
                        CostosOperacionales = cotizacion.GastoOperacional
                            .Select(c => new
                            {
                                tipo = c.TipoGastoId,
                                tipoName = c.TipoGasto.Descripcion,
                                valor = c.Valor,
                                cantidad = c.Cantidad
                            }).ToList(),
                        Files = cotizacion.ArchivoCotizacion
                                    .Select(f => new { IdFile = f.idArchivo })
                                    .ToList()
                    }

                };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object DisableFile(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);
                var file = db.Archivos.Find(id);
                if (file == null) return new { done = false, message = "Error, no existe archivo" };
                file.Habilitado = false;
                db.SaveChanges();
                return new { done = true, message = "Archivo deshabilitado" };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object UploadFile(HttpContext post)
    {
        try
        {
            var name = post.Request["name"];
            var uso = post.Request["uso"];
            string dirFullPath = HttpContext.Current.Server.MapPath("~/archivos_mail/");
            string[] files;
            files = System.IO.Directory.GetFiles(dirFullPath);
            string str_file = "";
            string str_extension = "";
            foreach (string s in post.Request.Files)
            {
                HttpPostedFile file = post.Request.Files[s];
                string fileName = file.FileName;
                string fileExtension = file.ContentType;

                if (!string.IsNullOrEmpty(fileName))
                {
                    str_extension = Path.GetExtension(fileName).Replace(".", "");
                    var random = Guid.NewGuid().ToString("N").Substring(0, 8);
                    str_file = random + "." + str_extension;
                    string pathToSave = dirFullPath + str_file;
                    file.SaveAs(pathToSave);
                }
            }

            using (var db = new CertelEntities())
            {
                var file = new Archivos
                {
                    Nombre = name,
                    URL = str_file,
                    Extension = str_extension,
                    Habilitado = true,
                    Uso = uso
                };
                db.Archivos.Add(file);
                db.SaveChanges();
            }

            return new { done = true, message = "Artículo subido correctamente", file = str_file };
        }
        catch (Exception ex)
        {
            return new { code = 0, message = ex.ToString() };
        }
    }
    private static object GetPriceLogisticCoste(HttpContext post)
    {
        try
        {
            var ciudad = int.Parse(post.Request["ciudad"]);
            var tipoGasto = int.Parse(post.Request["tipoGasto"]);
            var moneda = int.Parse(post.Request["moneda"]);
            using (var db = new CertelEntities())
            {
                var priceList = db.ListaPrecioCostosLogisticos
                                    .Where(w => w.CiudadId == ciudad)
                                    .Where(w => w.TipoGastoId == tipoGasto)
                                    .FirstOrDefault();
                if (priceList == null)
                    return new { done = false, price = 0 };

                double price = 0;
                var uf = db.Settings.Find("UF");
                var dolar = db.Settings.Find("Dolar");
                var euro = db.Settings.Find("Euro");
                if (uf == null && moneda == 3)
                {
                    return new { done = false, message = "No hay valor asociado a la UF. Intente salir del sistema e ingresar nuevamente." };
                }
                if (dolar == null && moneda == 2)
                {
                    return new { done = false, message = "No hay valor asociado al Dolar. Intente salir del sistema e ingresar nuevamente." };
                }
                if (euro == null && moneda == 4)
                {
                    return new { done = false, message = "No hay valor asociado al Euro. Intente salir del sistema e ingresar nuevamente." };
                }

                switch (moneda)
                {
                    case 1: // PESO
                    default:
                        price = (double)priceList.ValorCLP;
                        break;
                    case 2: // DOLAR
                        double DOLAR = double.Parse(dolar.Valor.Replace(".", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                        price = (double)(priceList.ValorCLP / DOLAR);
                        break;
                    case 3: // UF
                        double UF = double.Parse(uf.Valor.Replace(".", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                        price = (double)priceList.ValorCLP / UF;
                        break;
                    case 4: // EURO
                        double EURO = double.Parse(euro.Valor.Replace(".", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                        price = (double)(priceList.ValorCLP / EURO);
                        break;
                }


                return new { done = true, price = price };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString() };
        }
    }

    private static object GetPrice(HttpContext post)
    {
        try
        {
            var producto = int.Parse(post.Request["producto"]);
            var elevador = int.Parse(post.Request["elevador"]);
            var paradas = int.Parse(post.Request["paradas"]);
            var cantidad = int.Parse(post.Request["cantidad"]);
            var moneda = int.Parse(post.Request["moneda"]);

            if (producto == 0 || elevador == 0 || cantidad == 0)
            {
                return new { done = false, price = 0 };
            }
            using (var db = new CertelEntities())
            {
                var priceList = db.ListaPrecio
                                .Where(w => w.TipoElevadorId == elevador)
                                .Where(w => w.ProductoId == producto)
                                .Where(w => w.Paradas == paradas)
                                .Where(w => w.MasDe1Equipo == (cantidad > 1))
                                .FirstOrDefault();
                if (priceList == null)
                    return new { done = false, price = 0 };

                double price = 0;
                var uf = db.Settings.Find("UF");
                var dolar = db.Settings.Find("Dolar");
                var euro = db.Settings.Find("Euro");
                if (uf == null)
                {
                    return new { done = false, message = "No hay valor asociado a la UF. Intente salir del sistema e ingresar nuevamente." };
                }
                if (dolar == null && moneda == 2)
                {
                    return new { done = false, message = "No hay valor asociado al Dolar. Intente salir del sistema e ingresar nuevamente." };
                }
                if (euro == null && moneda == 4)
                {
                    return new { done = false, message = "No hay valor asociado al Euro. Intente salir del sistema e ingresar nuevamente." };
                }
                var ufEnPesos = double.Parse(uf.Valor.Replace(".", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                switch (moneda)
                {
                    case 1: // PESO
                    default:
                        price = (double)priceList.ValorUF * ufEnPesos;
                        break;
                    case 2: // DOLAR
                        double DOLAR = double.Parse(dolar.Valor.Replace(".", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                        price = (double)(priceList.ValorUF * ufEnPesos / DOLAR);
                        break;
                    case 3:
                        price = (double)priceList.ValorUF;
                        break;
                    case 4: // EURO
                        double EURO = double.Parse(euro.Valor.Replace(".", "").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                        price = (double)(priceList.ValorUF * ufEnPesos / EURO);
                        break;
                }

                return new { done = true, ufEnPesos, price = Math.Round(price, moneda == 1 ? 0 : 2) };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString() };
        }
    }
    private static object GetClientByRut(string rut)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var client = db.Cliente
                                .Where(w => w.Rut == rut)
                                .Select(s => new
                                {
                                    Rut = s.Rut,
                                    Nombre = s.Nombre,
                                    Direccion = s.Direccion,
                                    Telefono = s.TelefonoContacto,
                                    Giro = s.Giro,
                                    NombreContacto = s.NombreContacto,
                                    Email = s.EmailContacto
                                })
                                .FirstOrDefault();
                return new { done = true, client = client };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString() };
        }
    }

    private static object GetFiles(string type)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var files = db.Archivos
                                .Where(w => w.Uso == type)
                                .Where(w => w.Habilitado == true)
                                .Select(s => new
                                {
                                    Id = s.Id,
                                    Name = s.Nombre,
                                    Url = s.URL,
                                    Extension = s.Extension
                                })
                                .ToList();
                return new { done = true, message = "OK", data = files };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object GetGlobal(HttpContext post)
    {
        try
        {
            var config = new ConfigVariables();
            config.SetData();
            return new { done = true, data = config };
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.Message };
        }
    }
    private static object GetPDF(HttpContext post)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var id = int.Parse(post.Request["id"]);
                var cotizacion = db.Cotizacion.Find(id);
                if (cotizacion == null)
                {
                    return new { done = false, message = "Error: No existe cotización" };
                }
                var pdf = new PDFCotizacion(cotizacion);
                cotizacion.Filename = pdf.FileName;
                db.SaveChanges();
                return new { done = true, filename = pdf.FileName };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = ex.ToString() };
        }
    }
    private static object Save(HttpContext post)
    {
        try
        {

            using (var db = new CertelEntities())
            {
                var rut_cliente = post.Request["rut_cliente"];
                var nombre_cliente = post.Request["nombre_cliente"];
                var giro_cliente = post.Request["giro_cliente"];
                var direccion_cliente = post.Request["direccion_cliente"];
                var contacto_cliente = post.Request["contacto_cliente"];
                var email_contacto_cliente = post.Request["email_contacto_cliente"];
                var fono_contacto_cliente = post.Request["fono_contacto_cliente"];

                var existClient = db.Cliente.Where(w => w.Rut == rut_cliente).FirstOrDefault();
                if (existClient == null)
                {
                    existClient = new Cliente
                    {
                        Nombre = nombre_cliente,
                        Giro = giro_cliente,
                        Direccion = direccion_cliente,
                        NombreContacto = contacto_cliente,
                        EmailContacto = email_contacto_cliente,
                        TelefonoContacto = fono_contacto_cliente
                    };
                    db.Cliente.Add(existClient);

                }
                else
                {
                    existClient.Nombre = nombre_cliente;
                    existClient.Giro = giro_cliente;
                    existClient.Direccion = direccion_cliente;
                    existClient.NombreContacto = contacto_cliente;
                    existClient.EmailContacto = email_contacto_cliente;
                    existClient.TelefonoContacto = fono_contacto_cliente;

                }

                var fecha_doc = post.Request["fecha_doc"];
                var fuente_solicitud = post.Request["fuente_solicitud"];
                var forma_pago = post.Request["forma_pago"];
                var moneda = post.Request["moneda"];
                var vendedor = post.Request["vendedor"];
                var observacion = post.Request["observacion"];
                var nota = post.Request["nota"];
                var validez = post.Request["validez"];
                var usuario = post.Request["user"];
                var ciudad = int.Parse(post.Request["ciudad"]);
                var user = new Encriptacion(usuario, false).newText;
                var lastIt = 0;
                var last = db.Cotizacion.Where(w => w.IT != null).OrderByDescending(o => o.IT).FirstOrDefault();
                //var us = db.Usuario.Find(user);
                //if (us == null || us.Habilitado != true)
                //{
                //    return new { done = false, message = "El usuario indicado no existe, o ha sido deshabilitado" };
                //}
                //var mailMe = us.Email;
                //if (mailMe == null || us.PassMail == null)
                //{
                //    return new { done = false, message = "Error: Guarde una contraseña de su correo electrónico antes de guardar la cotización. Puede hacerlo desde el menú de la derecha" };
                //}
                if (last == null)
                {
                    var firstIt = db.Settings.Find("FIRST_IT");
                    lastIt = int.Parse(firstIt.Valor);
                }
                else
                {
                    lastIt = (int)last.IT + 1;
                }
                var cotizacion = new Cotizacion
                {
                    FechaDoc = DateTime.ParseExact(fecha_doc, "dd-MM-yyyy", null),
                    FuenteSolicitud = fuente_solicitud,
                    FormaDePagoId = int.Parse(forma_pago),
                    MonedaId = int.Parse(moneda),
                    Vendedor = vendedor,
                    Observacion = observacion,
                    Nota = nota,
                    ClienteId = existClient.ID,
                    EstadoId = 1,
                    Usuario = user.ToLower(),
                    FechaCreacion = DateTime.Now,
                    CiudadId = ciudad,
                    FechaValidez = DateTime.ParseExact(validez, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture),
                    Valor = 0,
                    Descuento = 0,
                    Recargo = 0,
                    Total = 0,
                    IT = lastIt
                };
                db.Cotizacion.Add(cotizacion);
                db.SaveChanges();
                var items = post.Request["items"];
                var deserializer = new JavaScriptSerializer();
                var itemsList = deserializer.Deserialize<List<Item_Cotizacion>>(items);
                double totalValor = 0;
                double totalDescuento = 0;
                double totalRecargo = 0;

                foreach (var i in itemsList)
                {
                    var existsProduct = db.Producto.Where(w => w.Codigo == i.Producto).FirstOrDefault();
                    if (existsProduct == null) continue;

                    var itemCotizacion = new ItemCotizacion
                    {
                        CotizacionId = cotizacion.Id,
                        ProductoId = existsProduct.Id,
                        Cantidad = i.Cantidad,
                        DescripcionEditada = i.Descripcion,
                        Descuento = i.Descuento,
                        ModoDescuento = i.ModoDescuento,
                        ValorDescuento = i.ValorDescuento,
                        Unitario = i.ValorUnitario,
                        Exento = true,
                        AlturaPisos = i.Altura,
                        InstalacionAno = i.AnoInstalacion,
                        EmpresaInstaladora = i.EmpresaInstaladora,
                        Ubicacion = i.Ubicacion

                    };

                    if(i.Marca != 0) {
                        itemCotizacion.MarcaID = i.Marca;
                    }
                    if(i.TipoElevador != 0) {
                        itemCotizacion.AparatoID = i.TipoElevador;
                    }
                    if(i.TipoFuncionamiento != 0) {
                        itemCotizacion.TipoFuncionamientoID = i.TipoFuncionamiento;
                    }
                    if(i.Uso != "0") {
                        itemCotizacion.Uso = i.Uso;
                    }
                    db.ItemCotizacion.Add(itemCotizacion);
                    db.SaveChanges();
                    totalValor += i.ValorUnitario * i.Cantidad;
                    totalDescuento += i.ValorDescuento * i.Cantidad;

                }

                var gastos = post.Request["gastos"];
                var gastosList = deserializer.Deserialize<List<CostoOperacional>>(gastos);

                foreach (var g in gastosList)
                {
                    var gasto = new GastoOperacional
                    {
                        Cantidad = g.cantidad,
                        CotizacionId = cotizacion.Id,
                        TipoGastoId = g.tipo,
                        Valor = g.valor
                    };
                    db.GastoOperacional.Add(gasto);
                    totalRecargo += g.valor * g.cantidad;
                }


                var files = post.Request["files[]"];
                if (files != null)
                {
                    if (files.Contains(","))
                    {
                        var array = files.Split(',');
                        foreach (var a in array)
                        {
                            var id = int.Parse(a);
                            var archivo = new ArchivoCotizacion
                            {
                                idCotizacion = cotizacion.Id,
                                idArchivo = id
                            };
                            db.ArchivoCotizacion.Add(archivo);
                        }
                    }
                    else
                    {
                        var id = int.Parse(files);
                        var archivo = new ArchivoCotizacion
                        {
                            idCotizacion = cotizacion.Id,
                            idArchivo = id
                        };
                        db.ArchivoCotizacion.Add(archivo);
                    }
                }
                var hito = new Hito
                {
                    CotizacionId = cotizacion.Id,
                    TipoHitoId = 0,
                    Fecha = DateTime.Now,
                    Observacion = "Cotización creada"
                };
                db.Hito.Add(hito);

                // TOTALES
                cotizacion.Valor = totalValor;
                cotizacion.Descuento = totalDescuento;
                cotizacion.Recargo = totalRecargo;
                cotizacion.Total = totalValor - totalDescuento + totalRecargo;

                db.SaveChanges();

                //var sendMail = bool.Parse(post.Request["sendMail"]);
                //bool sentMail = false;

                //if (sendMail)
                //{
                //    var settings = new ConfigVariables();
                //    settings.SetData();
                //    var asunto = post.Request["asunto"];
                //    var from = us.Email;
                //    var passFrom = new Encriptacion(us.PassMail, false).newText;
                //    var correo = post.Request["correo"];
                //    var para = post.Request["para[]"];
                //    var cc = post.Request["cc[]"];
                //    var cco = post.Request["cco[]"];
                //    var name = us.Nombre + " " + us.Apellido;
                //    var firma = us.Firma;
                //    sentMail = SendMail(para, cc, cco, asunto, from, passFrom, correo, cotizacion, mailMe, name, firma, true);
                //}

                return new {
                    done = true,
                    message = "Cotización agregada exitosamente",
                    id = cotizacion.Id
                };

            }
        }
        catch (Exception ex)
        {
            log.Error(ex);
            return new { done = false, message = ex.ToString() };
        }
    }

    private static object Reenviar(HttpContext post, bool nuevo)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var settings = new ConfigVariables();
                settings.SetData();
                var usuario = post.Request["user"];
                var asunto = post.Request["asunto"];
                var cotizacionId = int.Parse(post.Request["cotizacionId"]);
                var cotizacion = db.Cotizacion.Find(cotizacionId);
                if (cotizacion == null)
                    return new { done = false, message = "Error: cotización no existe" };
                var user = new Encriptacion(usuario, false).newText;
                var us = db.Usuario.Find(user);
                if (us == null || us.Habilitado != true)
                {
                    return new { done = false, message = "El usuario indicado no existe, o ha sido deshabilitado" };
                }
                var mailMe = us.Email;
                if (mailMe == null || us.PassMail == null)
                {
                    return new { done = false, message = "Error: Guarde una contraseña de su correo electrónico antes de guardar la cotización. Puede hacerlo desde el menú de la derecha" };
                }
                var from = us.Email;
                var passFrom = new Encriptacion(us.PassMail, false).newText;
                var correo = post.Request["correo"];
                var para = post.Request["para[]"];
                var cc = post.Request["cc[]"];
                var cco = post.Request["cco[]"];
                var name = us.Nombre + " " + us.Apellido;
                var firma = us.Firma;
                var hito = new Hito
                {
                    TipoHitoId = nuevo ? 1 : 18,
                    Observacion = nuevo ? "Cotización enviada" : "Cotización reenviada",
                    Fecha = DateTime.Now,
                    CotizacionId = cotizacionId
                };
                db.Hito.Add(hito);
                db.SaveChanges();
                var sentMail = SendMail(para, cc, cco, asunto, from, passFrom, correo, cotizacion, from, name, firma, true);
                return new { done = sentMail, message = sentMail ?  "Cotización reenviada nuevamente al cliente" : "No fue posible enviar el correo al cliente", sentMail };
            }
        }
        catch (Exception ex)
        {
            return new { done = false, message = "Ocurrió un error", ex = ex.ToString() };
        }
    }
    private static bool SendMail(string to, string cc, string cco, string asunto, string from, string passForm, string correo, Cotizacion cotizacion, string mailMe, string name, string firma, bool newcotizacion)
    {
        try
        {

            var config = new ConfigVariables();
            config.SetData();
            if (config == null) return false;
            var smtpServer = new SmtpClient("mail.certel.cl");
            smtpServer.Port = 25;
            smtpServer.Credentials = new System.Net.NetworkCredential(from, passForm);
            var mail = new MailMessage();
            mail.From = new MailAddress(from, string.Format("Certel - {0}", name));
            string[] tos = null;
            if (to.Contains(","))
            {
                tos = to.Split(',');
                foreach (var t in tos)
                {
                    if (IsValidEmail(t))
                        mail.To.Add(t); // para quien
                }
            }
            else
            {
                if (IsValidEmail(to))
                    mail.To.Add(to);
            }
            string[] ccs = null;
            if (!string.IsNullOrEmpty(cc) && cc.Contains(","))
            {
                ccs = cc.Split(',');
                foreach (var c in ccs)
                {
                    if (IsValidEmail(c))
                        mail.CC.Add(c); // con copia
                }
            }
            else
            {
                if (IsValidEmail(cc))
                    mail.CC.Add(cc);
            }


            string[] ccos = null;
            if (!string.IsNullOrEmpty(cco) && cco.Contains(","))
            {
                ccos = cco.Split(',');
                foreach (var c in ccos)
                {
                    if (IsValidEmail(c))
                        mail.Bcc.Add(c); // con copia oculta
                }
            }
            else
            {
                if (IsValidEmail(cco))
                    mail.Bcc.Add(cco); // Con copia oculta
            }
            if (ccs != null && !ccs.Contains(config.cotizacion_mail) && newcotizacion)
                mail.CC.Add(config.cotizacion_mail); // con copia
            if (tos != null && !tos.Contains(mailMe))
                mail.CC.Add(mailMe);
            mail.Subject = asunto;
            var body = GetHtmlToMail(correo, config.public_logo, cotizacion, firma, newcotizacion);
            mail.Body = body;
            mail.IsBodyHtml = true;
            mail.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");
            smtpServer.EnableSsl = true;

            smtpServer.Send(mail);
            return true;
        }
        catch (Exception ex)
        {
            return false;
            //return false;
        }
    }

    private static string GetHtmlToMail(string correo, string logo, Cotizacion cotizacion, string firma, bool newcotizacion)
    {
        var html = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//ES\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\" >";
        html += "<html xmlns=\"http://www.w3.org/1999/xhtml\">";
        html += "<p>";
        html += correo;
        html += "</p>";
        html += "<hr>";
        html += "<br>";
        if(newcotizacion) {
            var dir = "http://certificaciondeascensores.cl/archivoscotizacion.aspx?c=" + cotizacion.Id;
            html += "<a href='" + dir + "' target='_blank' style='font-size: 12pt; cursor: pointer, font-weigth: bold'>VEA SU COTIZACIÓN AQUÍ </a><br><br><hr>";
        }

        html += firma;
        html += "<img src='" + logo + "' width='50%' style='max-width:5cm'>";
        html += "<br>";

        return html;
    }

    public struct Item_Cotizacion
    {
        public string Producto { get; set; }
        public string Descripcion { get; set; }
        public int Cantidad { get; set; }
        public double ValorUnitario { get; set; }
        public double Descuento { get; set; }
        public string ModoDescuento { get; set; }
        public double ValorDescuento { get; set; }
        public double PrecioCliente { get; set; }
        public int? TipoElevador { get; set; }
        public int? TipoFuncionamiento { get; set; }
        public int? Marca { get; set; }
        public int AnoInstalacion { get; set; }
        public string EmpresaInstaladora { get; set; }
        public string Ubicacion { get; set; }
        public string Uso { get; set; }
        public int Altura { get; set; }
        public int Ciudad { get; set; }
    }

    public struct CostoOperacional
    {
        public int cantidad { get; set; }
        public int tipo { get; set; }
        public string tipoName { get; set; }
        public double valor { get; set; }
        public double total { get; set; }
    }
    private static object GetClientes(string sidx, string sord, int page, int rows, string name, string rut)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var list = db.Cliente
                                .Where(w => w.Nombre.Contains(name))
                                .Where(w => w.Rut.Contains(rut))
                                .Where(w => w.Habilitado);

                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    //.OrderBy(sidx + " " + sord)
                    .OrderByDescending(o => o.Habilitado)
                    .ThenBy(o => o.Nombre)
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
                               Rut = x.Rut,
                               Nombre = x.Nombre == null ? string.Empty : x.Nombre.ToUpper(),
                               Direccion = x.Direccion == null ? string.Empty : x.Direccion.ToUpper(),
                               Telefono = x.TelefonoContacto,
                               Email = x.EmailContacto == null ? string.Empty : x.EmailContacto.ToLower(),
                               Giro = x.Giro,
                               NombreContacto = x.NombreContacto
                           })
                           .ToList()
                };
                return grid;

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };

        }
    }

    private static object GetCotizaciones(string sidx, string sord, int page, int rows, string cliente, int estado, string inicio, string fin, bool alertF2)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var f1 = inicio == string.Empty ? new DateTime(2017, 1, 1) : DateTime.ParseExact(inicio, "dd-MM-yyyy", null);
                var f2 = fin == string.Empty ? DateTime.Now : DateTime.ParseExact(fin, "dd-MM-yyyy", null);
                var settingsDaysFase2 = db.Settings.Find("FASE2_ALERT_DAYS");
                var daysFase2 = settingsDaysFase2 == null ? 15 : int.Parse(settingsDaysFase2.Valor);
                var vencimiento = DateTime.Now.AddDays(daysFase2);
                var list = db.Cotizacion
                                .Where(w => w.Cliente.Nombre.Contains(cliente) || w.Cliente.Rut.Contains(cliente))
                                .Where(w => estado == 0 ? w.EstadoId != 5 : w.EstadoId == estado)
                                .Where(w => w.FechaDoc >= f1 && w.FechaDoc <= f2)
                                .Where(w => !alertF2 ? true : w.CreadaFase2 == true && w.AlertaFase2Enviada != true && (vencimiento > w.FechaCompromisoFase2.Value));

                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    .OrderBy(sidx + " " + sord)
                    //.OrderByDescending(o => o.FechaSolicitud)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize);


                var grid = new
                {
                    page = page,
                    records = totalRecords,
                    total = totalPages,
                    rows = result
                            .AsEnumerable()
                           .Select(x => new
                           {
                               Id = x.Id,
                               IT = x.IT,
                               ClienteId = x.ClienteId,
                               Cliente = x.Cliente.Nombre,
                               FechaDoc = x.FechaDoc.Value.ToString("dd-MM-yyyy"),
                               MailCliente = x.Cliente.EmailContacto,
                               NombreContacto = x.Cliente.NombreContacto,
                               Responsable = x.Usuario != null ? x.Usuario1.Nombre + " " + x.Usuario1.Apellido : string.Empty,
                               EstadoId = x.EstadoId,
                               Estado = x.EstadoCotizacion.Descripcion,
                               Moneda = x.Moneda.Descripcion,
                               Observacion = x.Observacion,
                               FechaCreacion = x.FechaCreacion.Value.ToString("dd-MM-yyyy"),
                               Valor = x.Valor,
                               Descuento = x.Descuento,
                               Recargo = x.Recargo,
                               Total = x.Total,
                               HasService = x.Servicio.Any(),
                               HasFase2 = x.CreadaFase2,
                               FechaFase2 = x.FechaCompromisoFase2,
                               AlertaFase2Enviada = x.AlertaFase2Enviada,
                               IsOk = x.Servicio.Any(a => !a.Inspeccion.Any(i => i.Calificacion == null)),
                               AlertFase2 = x.CreadaFase2 == true ?
                                            (
                                                (DateTime.Now.AddDays(daysFase2) > x.FechaCompromisoFase2.Value)
                                            )
                                            : false,
                                SentMail = x.Hito.Any(a => a.TipoHitoId == 1)
                           })
                           .ToList()
                };
                return grid;

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.ToString() };

        }
    }


    private static object GetProductos(string sidx, string sord, int page, int rows, string name)
    {
        try
        {
            using (var db = new CertelEntities())
            {
                var list = db.Producto
                                .Where(w => w.Codigo.Contains(name) || w.Descripcion.Contains(name));

                int pageSize = rows;
                int totalRecords = list.Count();
                int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);
                totalPages = totalPages == 0 ? 1 : totalPages;
                var result = list
                    //.OrderBy(sidx + " " + sord)
                    .OrderBy(o => o.Descripcion)
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
                               Id = x.Id,
                               Codigo = x.Codigo,
                               Descripcion = x.Descripcion,
                               Precio = x.PrecioReferencia,
                               Moneda = x.MonedaId == null ? "" : x.Moneda.Descripcion
                           })
                           .ToList()
                };
                return grid;

            }
        }
        catch (Exception ex)
        {
            return new { done = false, ex = ex.Message };

        }
    }
    public bool IsReusable
    {
        get
        {
            return true;
        }
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            if (string.IsNullOrEmpty(email))
                return false;
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }

}